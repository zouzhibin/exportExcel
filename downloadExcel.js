// 下载模板
function outputXlsxFile(XLSX, data, wscols, xlsxName) {
  var sheetNames = [];
  var sheetsList = {};
  for (var key in data) {
    sheetNames.push(key);
    var temp = DataExcel(data[key]);
    sheetsList[key] = XLSX.utils.aoa_to_sheet(temp);
    sheetsList[key]['!cols'] = wscols;
  }
  const wb = XLSX.utils.book_new();
  wb['SheetNames'] = sheetNames;
  wb['Sheets'] = sheetsList;
  XLSX.writeFile(wb, xlsxName + ".xlsx");
  //处理数据的函数
  function DataExcel(data) {
    var total = [];
    var temp = data.xlsxHeader;
    // var temp = [];
    total.push(temp);
    data.data.forEach(item => {
      var arr = [];
      data.xlsxAttribute.map((v) => {
        arr.push(item[v])
      })
      total.push(arr);
    })
    return total;
  }
}

function outputXlsxInArray(opt) {
  const wopts = { bookType: 'xlsx', bookSST: true, type: 'binary' }
  // 这里的this 是传入过来的 XLSX ，因为外部进行了bind绑定函数
  let workbook = this.utils.book_new() // 创建一个工作表 workbook
  const sheets = opt['SheetNames'] // 获取sheet表名
  const sheetsConfig = []
  // SheetNames 里面保存了所有的sheet名字
  workbook['SheetNames'] = sheets
  sheets.map((v, i) => {
    const sheet = opt['Sheets'][i]
    const { merger, headers, config, ahead } = sheet
    // 把 headers 数组插入到 data的首部，相当于其实就是excel的标题头部
    sheet.data.unshift(headers)
    if (ahead && (ahead.constructor === Array)) sheet.data = [...ahead, ...sheet.data]
    if (merger) sheet.data.unshift(merger)
    // sheet.data 其实是一个二维数组，
    const sheetConfig = this.utils.aoa_to_sheet(sheet.data) // 将JS数据数组的数组转换为工作表

    // 这里的作用是把 !col 的数据绑定到s sheetConfig上 
    if (config) {
      const keys = Object.keys(config)
      keys.map(v => {
        const a = sheetConfig[v]
        if (!a) sheetConfig[v] = config[v]
        else sheetConfig[v] = Object.assign(a, config[v])
      })
    }
    console.log('ddd',sheetConfig)
    sheetsConfig[v] = sheetConfig
  })
  workbook['Sheets'] = sheetsConfig
  console.log('dddrr',sheetsConfig)
  const xlsxName = `${opt.xslx || '报表'}.xlsx` // 导出的excel名字
  var wbout = this.write(workbook, wopts) // 写入工作簿workbook
  // //创建二进制对象写入转换好的字节流
  var blob = new Blob([s2ab(wbout)], {type: ''}) // 创建一个 Blob 对象


  function saveAs (obj, fileName) {
    var tmpa = document.createElement('a')
    tmpa.download = fileName || '下载.xlsx'
    tmpa.href = URL.createObjectURL(obj) // 根据传入的参数创建一个指向该参数对象的URL
    tmpa.click()
    setTimeout(function () {
      URL.revokeObjectURL(obj)
    }, 100)
  }
  
  // 转换为字符流
  function s2ab (s) {
    var buf
    if (typeof ArrayBuffer !== 'undefined') {
      buf = new ArrayBuffer(s.length)
      var view = new Uint8Array(buf)
      for (let i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF
      return buf
    } else {
      buf = new Array(s.length)
      for (let i = 0; i !== s.length; ++i) buf[i] = s.charCodeAt(i) & 0xFF
      return buf
    }
  }

  saveAs(blob, xlsxName)
}


