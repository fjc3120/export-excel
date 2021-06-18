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
 * exportData为一个对象，包含三个元素: Array-dataList(表格数据)  Object-option(配置信息) Array-columnsList(列头信息)
 * let columnsList = [
      { name: '第一列name', code : '第一列code'},
      { name: '第二列name', code : '第二列code'},
      ...
    ];
    let dataList = [
        {'第一列code': '第一列value', '第一列code': '第一列value'}, // 第一行数据
        {'第二列code': '第二列value', '第二列code': '第二列value'}, // 第二行数据
        ...
    ];
    let option = {
        title: 'excel表单' || '未命名', // excel文件标题名
        width: 150, // 单元格宽度
        fontSize: 12, // 字体大小
        fontBold: true, // 文字是否加粗
        fillColor: 'B7DEEA', // 单元格背景色(颜色编码没有#)
        alignmentVertical: 'center', // 垂直
        alignmentHorizontal: 'center', // 水平
    };
 */
