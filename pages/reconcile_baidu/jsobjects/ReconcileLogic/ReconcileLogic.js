export default {
    excelData: [],
    matchResults: [],
    dbOrderMap: {},

    parseExcel: async function() {
      const file = FilePicker1.files[0];
      if (!file) {
        showAlert('请先选择Excel文件', 'warning');
        return;
      }

      const base64Data = file.data.split(',')[1];
      const workbook = XLSX.read(base64Data, { type: 'base64' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(sheet);

      this.excelData = data;
      showAlert('解析成功，共 ' + data.length + ' 条，正在匹配...', 'info');

      const dbOrders = await QueryBatchCheck.run();

      this.dbOrderMap = {};
      for (let i = 0; i < dbOrders.length; i++) {
        const order = dbOrders[i];
        this.dbOrderMap[order['百度单号']] = order;
      }

      const results = [];
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const orderNo = String(row['订单编号'] || '');
        const supplier = row['所属供应商'] || '';
        const dbOrder = this.dbOrderMap[orderNo];

        let status = '正常';
        let reason = '';

        if (!dbOrder) {
          status = '异常';
          reason = '单号不存在';
        } else if (dbOrder['客户'] !== supplier) {
          status = '异常';
          reason = '供应商不匹配';
        } else if (dbOrder['百度对账批次']) {
          status = '已对账';
          reason = dbOrder['百度对账批次'];
        }

        results.push({
          订单编号: orderNo,
          所属供应商: supplier,
          商品采购金额: row['商品采购金额'] || 0,
          物流采购金额: row['物流采购金额'] || 0,
          订单状态: row['订单状态'] || '',
          是否退款: row['是否全额退款'] || '否',
          匹配状态: status,
          异常原因: reason
        });
      }

      this.matchResults = results;
      showAlert('匹配完成!', 'success');
      return results;
    },

    doReconcile: async function() {
      const normalRecords = this.matchResults.filter(function(r) {
        return r['匹配状态'] === '正常';
      });

      if (normalRecords.length === 0) {
        showAlert('没有可导入的正常记录', 'warning');
        return;
      }

      const batchNo = 'BD' + moment().format('YYYYMMDDHHmmss');
      let totalAmount = 0;

      showAlert('正在导入 ' + normalRecords.length + ' 条...', 'info');

      for (let i = 0; i < normalRecords.length; i++) {
        const record = normalRecords[i];
        await QueryUpdateOrder.run({
          orderNo: record['订单编号'],
          productAmount: record['商品采购金额'] || 0,
          logisticsAmount: record['物流采购金额'] || 0,
          orderStatus: record['订单状态'] || '',
          isRefund: record['是否退款'] || '否',
          batchNo: batchNo
        });
        totalAmount += (record['商品采购金额'] || 0) + (record['物流采购金额'] || 0);
      }

      await QueryCreateBatch.run({
        batchNo: batchNo,
        type: '百度',
        count: normalRecords.length,
        total: totalAmount
      });

      showAlert('完成! 批次:' + batchNo + ' 导入:' + normalRecords.length + '条', 'success');
    },

    getStats: function() {
      const total = this.matchResults.length;
      const normal = this.matchResults.filter(function(r) { return r['匹配状态'] === '正常'; }).length;
      const error = this.matchResults.filter(function(r) { return r['匹配状态'] === '异常'; }).length;
      const done = this.matchResults.filter(function(r) { return r['匹配状态'] === '已对账'; }).length;
      return { total: total, normal: normal, error: error, done: done };
    }
  }