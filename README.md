# exportExcel
###  js-xlsx提供的接口非常清晰主要分为两类:
#### xlsx对象本身提供的功能
  - 解析数据
  - 导出数据
#### utils工具类
  - 将数据添加到数据表对象上
  - 将二维数组以及符合格式的对象或者HTML转为工作表对象
  - 将工作簿转为另外一种数据格式
  - 行,列,范围之间的转码和解码
  - 工作簿操作
  - 单元格操作

- https://www.npmjs.com/package/xlsx-style



- 示例问题 
参考 https://www.jianshu.com/p/877631e7e411
目前：文档里的./xlsx是整合后的代码

- 我对比了xlsx-style和xlsx的源码，
发现两个库很多的方法和代码都是一样的
，xlsx-style之所以能够设置样式，是因为它多了一个styleBuilder

- 而xlsx-style不能够设置行高，
是因为在write_ws_xml-data里没有对rows进行处理，而在xlsx库里可以清楚的看到它对!rows的处理

所以解决方法就必须去改库（我认为），思路也有2中：
（1）在xlsx-style加xlsx里设置行高的代码
（2）在xlsx里加入xlsx-style里设置样式的代码


参考 https://segmentfault.com/a/1190000019700368?utm_source=tag-newest
