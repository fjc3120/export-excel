# export-excel
前端导出excel

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
 * @LastEditTime: 2022-02-09 16:57:05
 * @Copyright: Copyright (c) 2018, Hand
 * 前端导出组件的配置信息option相关参考请看下文
 * https://www.cnblogs.com/liuxianan/p/js-excel.html
 * obj为一个对象，包含六个元素: Array-dataList(表格数据/必填)  Object-option(配置信息/选填) Array-columnsList(列头信息/必填) String-title(文件名/标题名/选填)优先级最高 Boolean-columnHeadMerge(开启双行列头合并、选填) Object-columnHeader(双核列头合并的第一行列头信息/columnHeadMerge开启时为必填,未开启时为不填)
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
    let columnHeadMerge = true; // 是否开启双行列头合并功能
    let columnHeader = { // 第一行信息(key为第二行的code，value为第一行的name),数量必须与columnsList一致
        companyName: '地市',
        periodName: '期间',
        ...
    };
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
 */
