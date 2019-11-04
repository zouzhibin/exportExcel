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


/**
 * 
js-xlsx提供的接口非常清晰主要分为两类:
xlsx对象本身提供的功能
  - 解析数据
  - 导出数据
utils工具类
  - 将数据添加到数据表对象上
  - 将二维数组以及符合格式的对象或者HTML转为工作表对象
  - 将工作簿转为另外一种数据格式
  - 行,列,范围之间的转码和解码
  - 工作簿操作
  - 单元格操作
 * @param {*} opt 
 */
function outputXlsxInArray(opt) {
  // //这里的数据是用来定义导出的格式类型
  const wopts = {  
    cellStyles: true,
    bookType: 'xlsx',  // 要生成的文件类型
    bookSST: true, //是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
    type: 'binary' }
  // 这里的this 是传入过来的 XLSX ，因为外部进行了bind绑定函数
  let workbook = this.utils.book_new() // 创建一个新的工作簿对象: workbook
  const sheets = opt['SheetNames'] // 获取sheet表名
  const sheetsConfig = []
  // SheetNames 里面保存了所有的sheet名字
  workbook['SheetNames'] = sheets
  sheets.map((v, i) => {
    const sheet = opt['Sheets'][i]

    /**
     * data["!merges"] = [{
            s: {//s为开始
                c: 1,//开始列
                r: 0//可以看成开始行,实际是取值范围
            },
            e: {//e结束
                c: 4,//结束列
                r: 0//结束行
            }
        }];
     */
    const { merges, headers, config, ahead } = sheet
    // 把 headers 数组插入到 data的首部，相当于其实就是excel的标题头部
    sheet.data.unshift(headers)
    // sheet.data.unshift(['录用名单'])
    if (ahead && (ahead.constructor === Array)) sheet.data = [...ahead, ...sheet.data]
    // sheet.data 其实是一个二维数组，
    // 二维数组的关系非常容易理解,数组中的每一个数组代表一行.
    // 将JS数组的数组（[ [...],[...],[...] ]）转换为工作表
    const sheetConfig = this.utils.aoa_to_sheet(sheet.data) // 将JS数据数组的数组转换为工作表 =>  使用二维数组创建一个工作表对象 
    
    // if (merges) sheetConfig['!merges'] = merges // 合并单元格

    sheetConfig["B1"].s = {  font: {
      name: '宋体',
      sz: 16,
      color: {rgb: "#FFFF0000"},
      bold: false,
      italic: false,
      underline: false
    },
    alignment: {
      horizontal: "center" ,
      vertical: "center"
    }};//<====设置xlsx单元格样式

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
    var tmpa = document.createElement('a') //绑定a标签
    tmpa.download = fileName || '下载.xlsx' 
    tmpa.href = URL.createObjectURL(obj) // 根据传入的参数创建一个指向该参数对象的URL //绑定a标签
    tmpa.click() // 模拟点击实现下载
    setTimeout(function () {
      URL.revokeObjectURL(obj) //用URL.revokeObjectURL()来释放这个object URL
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


