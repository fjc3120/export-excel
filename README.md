# export-excel
前端导出excel

/*
 * @Descripttion: 前端导出组件
 * @Author: 小鸟游露露
 * @Date: 2021-06-02 10:28:40
 * @LastEditTime: 2021-06-02 10:29:13
 * @Copyright: Copyright (c) 2018, Hand
 * 前端导出组件的配置信息option相关参考请看下文
 * https://www.cnblogs.com/liuxianan/p/js-excel.html
 * obj为一个对象，包含四个元素: Array-dataList(表格数据)  Object-option(配置信息) Array-columnsList(列头信息) String-title(文件名/标题名)优先级最高
 * let columnsList = [
      { name: '第一列name', code : '第一列code'},
      { name: '第二列name', code : '第二列code'},
      ...
    ];
    let dataList = [
        {'第一行第一列code': '第一行第一列value', '第一行第二列code': '第一行第二列value'}, // 第一行数据
        {'第二行第一列code': '第二行第一列value', '第二行第二列code': '第二行第二列value'}, // 第二行数据
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
 */
