# export-excel
前端导出excel

***1.1.33-alpha.3版本***

- 修复title各个sheet页的顶部标题未能正确展示的问题

***1.1.33-alpha.2版本***

- 修改option变为数组格式，支持各个sheet页使用不同的option

***1.1.33-alpha.1版本***

- 修改大部分属性强制使用多sheet页数据格式，支持多sheet页导出。目前支持columnsList、dataList、title作为数组。
- 新增sheetName（sheet页名字）和fileName（文件名）属性

***1.1.32版本***

- 删除c7n的notification消息弹框组件，改用window自带API，缩减体积；

***1.1.31版本***

- 修复开启表头合并后或dataList有多余字段时需要手动对dataList排序对象key的问题；

***1.1.30版本***

- 新增数字格式转化为携带千分位符的字符串功能；
- 修复dataList中存在字段缺失的情况时，该单元格展示undefined的问题。改为根据columnsList自动补全，展示空白单元格；
- 修复dataList中存在字段为null或undefined时，该单元格展示null或undefined的问题。改为展示空白单元格；

***1.1.29版本***

- 修复使用columnHeader时，表头最后一行的样式不能正常显示的问题；

***1.1.28版本***

- 修改表头的单元格内容超出宽度后自动换行；
- 新增columnHeaderGroup属性，支持任意行数的组合列作为表格的表头，注意columnHeaderGroup会比columnHeader的优先级更高，columnHeader被默认为columnHeaderGroup[0]

***1.1.27版本***

- 修复表格数据有null或undefined等时，单元格未展示空白而是显示'null'等问题；

***1.1.26版本***

- 修复使用双行猎头合并功能时列头显示错误的问题；

***1.1.25版本***

- 修复keyMap不存在报错问题；


***1.1.24版本***

- README文档更新

***1.1.23版本***

- 新增是否开启双行猎头合并功能——columnHeadMerge
- 新增第一行列头映射表——columnHeader

***1.1.22版本***

- 新增设置标题字体大小功能——fontSizeTitle

- 新增设置表格数据字体大小功能——fontSizeList

- 修改fontSize的功能为设置列头字体大小

- 修改fontBold的功能为设置列头字体是否加粗

- 修改fillColor的功能为设置列头单元格背景色

- 修复样式option未设置时，默认样式未生效问题

***1.1.21版本***

- 新增以列为维度的单元格宽度修改——option.widthColums

- 新增单元格合并功能——option.merges

***1.1.20版本***

- 新增对特定的批量的单元格定制化样式——同组单元格的样式相同，常用于对某个或某几个单元格的样式修改——option.styleGroup

- 新增对特定的某几行单元格定制化样式——同组单元格的样式相同——option.styleRow

/*
 * @Descripttion: 前端导出组件
 * @Author: 小鸟游露露
 * @Date: 2021-06-02 10:28:40
 * @LastEditTime: 2023-04-17 10:15:13
 * @Copyright: Copyright (c) 2018, Hand
 * 实参obj为一个对象，包含以下元素: Array-dataList(表格数据)  Object-option(配置信息) Array-columnsList(列头信息) Boolean-columnHeadMerge(是否开启列头合并功能) Array-columnsList(最后一行表列头映射表)  Array-columnHeader(第一行表头信息) Array-columnHeaderGroup(第一行至倒数第二行表头信息)
 * https://open.hand-china.com/community/detail/605115219152343040
 * https://www.cnblogs.com/liuxianan/p/js-excel.html
 * 【注意】当前版本为测试版，必须使用sheet页写法（允许仅一页）以下备注为单个sheet内部数据


 * let columnsList = [ // columnHeadMerge 开启后做为第二行列头映射表
      { name: '第一列name', code : '第一列code'},
      { name: '第二列name', code : '第二列code'},
      ...
    ];
    let dataList = [
        {'第一行第一列code': '第一行第一列value', '第一行第二列code': '第一行第二列value'}, // 第一行数据
        {'第二行第一列code': '第二行第一列value', '第二行第二列code': '第二行第二列value'}, // 第二行数据
        ...
    ];
    let title = 'excel表单' || '未命名';
    let columnHeadMerge = true; // 是否开启列头合并功能
    let columnHeader = { // 第一行表头信息(key为第二行的code，value为第一行的name),属性数量必须与columnsList一致
        companyName: '地市',
        periodName: '期间',
        ...
    };
    let columnHeaderGroup = [ // 优先级高于columnHeader，第一行至倒数第二行列头信息(key为最后一行列头的数据列的code，value为对应行的name),属性数量必须与columnsList一致
      {
        companyName: '地市',
        periodName: '期间',
        ...
      },
      ...
    ];
    let option = {
        title: 'excel表单' || '未命名', // excel文件标题名
        width: 150, // 单元格宽度
        fontSize: 12, // 字体大小-列头
        fontSizeTitle: 14, // 字体大小-标题
        fontSizeList: 10, // 字体大小-表格
        fontBold: true, // 列头文字是否加粗
        fillColor: 'B7DEEA', // 列头单元格背景色(颜色编码没有#)
        alignmentVertical: 'center', // 垂直
        alignmentHorizontal: 'center', // 水平
        styleGroup: [ // 批量单元格自定制样式
        {
          cells: ['A3', 'A11', 'A19', 'A27'],
          style: {
            font: { sz: 10, color: { rgb: '4fd2db' } },
            alignment: { vertical: 'center', horizontal: 'bottom', wrapText: 'true' },
          },
        },
        {
          cells: ['B5'],
          style: {
            font: { sz: 10, color: { rgb: 'cccccc' } },
            alignment: { vertical: 'center', horizontal: 'bottom', wrapText: 'true' },
          },
        },
      ],
      styleRow: [ // 批量单元格——以一行为维度，自定制样式
        {
          cells: ['3', '5'],
          style: {
            font: { sz: 10, color: { rgb: 'cccccc' } },
            alignment: { vertical: 'center', horizontal: 'bottom', wrapText: 'true' },
          },
        },
        {
          cells: ['6'],
          style: {
            font: { sz: 14, color: { rgb: 'cccccc' } },
            alignment: { vertical: 'center', horizontal: 'bottom', wrapText: 'true' },
          },
        },
      ],
      widthColums: [ // 以列为维度设置单元格宽度
        {
          cells: [3, 5],
          width: 100,
        },
        {
          cells: [1],
          width: 200,
        },
      ],
      merges: [ // 单元格合并，s-e 代表区域 c-r 代表列-行的索引
        {
          s: { c: 0, r: 2 },
          e: { c: 0, r: 5 },
        },
        {
          s: { c: 0, r: 10 },
          e: { c: 3, r: 10 },
        },
      ],
    };
    let title = 'excel表单' || '未命名';

    本次sheet页修改后的数据传输格式
    columnsList、dataList、title、option都变为在原基础上再加一层数组包裹
    其中title功能修改，从原本的标题+文件名 改为 各个sheet的标题
    新增fileName作为文件名
    新增sheetName作为各个sheet页的名字(左下角)
    obj.columnsList = [columnsList1, columnsList2];
    obj.option = [option1, option2];
    obj.dataList = [dataList1, dataList2];
    obj.title = [customFileName1, customFileName2];
    obj.sheetName = [sheetName1, sheetName2];
    obj.fileName = fileName;
 */
