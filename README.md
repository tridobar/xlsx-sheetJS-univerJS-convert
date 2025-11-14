# xlsx-sheetJS-univerJS-convert
将xlsx/sheetJS读取的文件转换为univer sheet的IWorkbookData类型

xlsx.js或xlsx-js-style读取的文件好像都没有样式，所以这里也**不支持样式转换**

需要导入样式，可以使用excelJS，并使用[excel-univer-convert](https://github.com/tridobar/excelJS-univerJS-convert)

# TS代码
```TypeScript
// 读取文件
new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = e.target.result;
      const workbook = XLSX.read(data, { type: 'binary' });
      resolve(this.convertWorkbookToJson(workbook, {
        id: new Date().getTime().toString(),
        name: file.name,
        appVersion: '',
        locale: UniverCore.LocaleType.ZH_CN,
      }));
    } catch (error) {
      reject(error);
    }
  };
  reader.onerror = (error) => reject(error);
  reader.readAsBinaryString(file);
})

// xlsx json 转 univer json
convertWorkbookToJson = (workbook, workbookConfig: any = {}) => {
  const sheets = {};
  const sheetOrder = [];
  const utils = XLSX.utils;
  workbook.SheetNames.forEach((sheetName, sheetIndex) => {
    const worksheet = workbook.Sheets[sheetName];
    const jsonSheet = utils.sheet_to_json(worksheet, { header: 1 });
    const cellData = {};
    let maxColumnCount = 0;
    jsonSheet.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        if (cell !== null && cell !== undefined && cell !== "") {
          if (!cellData[rowIndex]) {
            cellData[rowIndex] = [];
          }
          // 赋值
          cellData[rowIndex][colIndex] = { v: cell };
          // 公式
          const workCell = worksheet[utils.encode_cell({ c: colIndex, r: rowIndex })];
          cellData[rowIndex][colIndex].f = workCell.f;
          // 计算最大列数
          if (colIndex + 1 > maxColumnCount) {
            maxColumnCount = colIndex + 1;
          }
        }
      });
    });
    const sheetId = `sheet_${sheetIndex}`;
    sheets[sheetId] = {
      id: sheetId,
      name: sheetName,
      rowCount: jsonSheet.length + 50,
      columnCount: maxColumnCount + 50,
      zoomRatio: 1,
      cellData: cellData,
      showGridlines: 1,
      mergeData: [],
    };
    // 处理合并单元格
    worksheet['!merges']?.forEach(merge => {
      sheets[sheetId].mergeData.push({
        startRow: merge.s.r,
        startColumn: merge.s.c,
        endRow: merge.e.r,
        endColumn: merge.e.c,
        rangeType: 0,
        unitId: sheetId,
        sheetId: workbookConfig.id,
      });
    });
    // 样式好像导入不了

    sheetOrder.push(sheetId);
  });
  return { ...workbookConfig, sheetOrder: sheetOrder, sheets: sheets };
};

// univer json 转 xlsx json
reverseConvertJsonToWorkbook = (jsonData) => {
  const workbook = { SheetNames: [], Sheets: {} };
  const utils = XLSX.utils;

  // 遍历所有 sheet（按原顺序）
  jsonData.sheetOrder.forEach(sheetId => {
    const sheetInfo = jsonData.sheets[sheetId];
    const sheetName = sheetInfo.name;
    const cellData = sheetInfo.cellData;

    // 创建新的工作表对象
    const newSheet = {
      '!ref': "A1:" + utils.encode_cell({ c: sheetInfo.columnCount, r: sheetInfo.rowCount }),
      '!rows': [],
      '!cols': [],
      '!merges': [],
    };

    let maxColumnCount = 0;
    // 遍历单元格数据
    for (const rowKey in cellData) {
      const rowIndex = parseInt(rowKey);
      const rowCells = cellData[rowKey];
      maxColumnCount = Math.max(maxColumnCount, rowCells.length);

      for (const colKey in rowCells) {
        const colIndex = parseInt(colKey);
        const cell = rowCells[colKey];

        if (!cell) continue;

        // 生成单元格地址
        const cellAddress = utils.encode_cell({ r: rowIndex, c: colIndex });

        // 重建单元格对象
        newSheet[cellAddress] = {
          v: cell.v, // 原始值
          t: cell.t, // 类型
        };

        // 恢复公式
        if (cell.f) newSheet[cellAddress].f = cell.f;
      }
      // 设置行高
      if (sheetInfo.rowData[rowKey]) {
        newSheet['!rows'].push({ hpx: sheetInfo.rowData[rowKey].h });
      } else {
        newSheet['!rows'].push({});
      }
    }
    // 设置列宽
    for (let i = 0; i < maxColumnCount; i++) {
      if (sheetInfo.columnData[i]) {
        newSheet['!cols'].push({ wpx: sheetInfo.columnData[i].w });
      } else {
        newSheet['!cols'].push({});
      }
    }
    // 遍历合并行
    newSheet['!merges'] = sheetInfo.mergeData.map(merge => ({
      s: { r: merge.startRow, c: merge.startColumn },
      e: { r: merge.endRow, c: merge.endColumn },
    }));

    // 将工作表添加到 workbook
    workbook.SheetNames.push(sheetName);
    workbook.Sheets[sheetName] = newSheet;
  });

  return workbook;
};
