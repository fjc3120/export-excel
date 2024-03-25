/*
 * @Descripttion: 前端导出组件
 * @Author: 小鸟游露露
 * @Date: 2021-06-02 10:28:40
 * @LastEditTime: 2023-04-17 10:15:13
 * @Copyright: Copyright (c) 2018, Hand
 * 实参obj为一个对象，包含以下元素: Array-dataList(表格数据)  Object-option(配置信息) Array-columnsList(列头信息) Boolean-columnHeadMerge(是否开启列头合并功能) Array-columnsList(最后一行表列头映射表)  Array-columnHeader(第一行表头信息) Array-columnHeaderGroup(第一行至倒数第二行表头信息)
 * https://open.hand-china.com/community/detail/605115219152343040
 * https://www.cnblogs.com/liuxianan/p/js-excel.html
 * 【注意】当前版本为测试版，必须使用sheet页写法（允许仅一页），但样式和配置所有sheet共享（包括单元格样式、option、列合并、行合并等），以下备注为单个sheet内部数据
 *
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
    columnsList、dataList、title都变为在原基础上再加一层数组包裹
    其中title功能修改，从原本的标题+文件名 改为 各个sheet的标题
    新增fileName作为文件名
    新增sheetName作为各个sheet页的名字(左下角)
    obj.columnsList = [columnsList1, columnsList2];
    obj.dataList = [dataList1, dataList2];
    obj.title = [customFileName1, customFileName2];
    obj.sheetName = [sheetName1, sheetName2];
    obj.fileName = fileName;
 */
    import XLSX from './xlsx-style';
    function ExportExcel(obj={}) {
      // 点击导出按钮
      function handleExportData() {
        if (Object.keys(obj).length) {
          let {
            dataList = [],
            option = {},
            columnsList = [],
            title = [],
            columnHeadMerge = false,
            columnHeader = {},
            columnHeaderGroup = [],
            sheetName = ['sheet'],
            fileName = '未命名',
          } = obj;
          option.title = option.title || '未命名';
          option.width= option.width || 150;
          option.fontSize= option.fontSize || 12;
          option.fontSizeTitle= option.fontSizeTitle || 14;
          option.fontSizeList= option.fontSizeList || 10;
          option.fontBold= option.fontBold || true;
          option.fillColor= option.fillColor || 'B7DEEA';
          option.alignmentVertical= option.alignmentVertical || 'center';
          option.alignmentHorizontal= option.alignmentHorizontal || 'center';
          option.styleGroup= option.styleGroup || [];
          option.styleRow= option.styleRow || [];
          option.widthColums= option.widthColums || [];
          option.merges= option.merges || [];
          const optionSheetArr = [];
          const titleSheetArr = [];
          const colListSheetArr = [];
          const columnsListSheetArr = [];
          const dataAllSheetArr = [];
          const dataSheetArr = [];
          const dataExportSheetArr = [];
          const promises = dataList.map((itemDataList, indexDataList) => {
            return new Promise((resolve) => {
              // 动态创建多个变量
              optionSheetArr[indexDataList] = option;
              titleSheetArr[indexDataList] = title[indexDataList];
              colListSheetArr[indexDataList] = [];
              columnsListSheetArr[indexDataList] = columnsList[indexDataList];
              dataAllSheetArr[indexDataList] = JSON.parse(JSON.stringify(itemDataList));
              dataSheetArr[indexDataList] = [];
              dataExportSheetArr[indexDataList] = [];
              // 处理各个sheet的页面顶部标题
              if (title[indexDataList]) {
                optionSheetArr[indexDataList].title = title[indexDataList];
              }
              columnsListSheetArr[indexDataList].forEach(item => {
                colListSheetArr[indexDataList].push(item.code);
              });
    
              // 根据columnsList补全dataList缺失的字段
              for (let i = 0; i < dataAllSheetArr[indexDataList].length; i++) {
                let row = dataAllSheetArr[indexDataList][i];
                for (let j = 0; j < colListSheetArr[indexDataList].length; j++) {
                  let key = colListSheetArr[indexDataList][j];
                  if (!Object.prototype.hasOwnProperty.call(row, key)) {
                    row[key] = '';
                  }
                }
              }
    
              // 对dataList每行数据的key进行排序(不重写会导致列的顺序错乱)
              dataAllSheetArr[indexDataList].forEach(item => {
                const objItem = {};
                colListSheetArr[indexDataList].forEach(i => {
                  objItem[i] = item[i];
                });
                dataSheetArr[indexDataList].push(objItem);
              });
              dataExportSheetArr[indexDataList] = dataSheetArr[indexDataList].map((item) => {
                let newItem = {...item};
                for (let key in newItem) {
                  if (Object.prototype.hasOwnProperty.call(newItem, key)) {
                    if (typeof newItem[key] === 'number') { // 数字格式转化为带千分位符的字符串格式
                      let num = newItem[key];
                      let str = num.toLocaleString();
                      newItem[key] = str;
                    } else if (newItem[key] === null || newItem[key] === undefined) { // 赋予空值
                      newItem[key] = '';
                    }
                  }
                }
                return newItem;
              });
              resolve(); // 一旦操作完成，就resolve这个Promise
            });
          });
          Promise.all(promises).then(() => {
            exportRun(dataExportSheetArr, optionSheetArr, columnsListSheetArr, columnHeadMerge, columnHeader, columnHeaderGroup, sheetName, fileName); // 数据-配置信息-列头信息-是否开启列头合并-列头第一行映射表-sheet名列表-文件名
          });
        }
      }
    
      // 执行导出方法
      function exportRun(data, option, columns, columnHeadMerge, columnHeader, columnHeaderGroup, sheetName, fileName) {
        const dataJsonSheetArr = [];
        const columnsLenSheetArr = [];
        const promises = data.map((itemDataList, indexDataList) => {
          return new Promise((resolve) => {  
            dataJsonSheetArr[indexDataList] = [];
            columnsLenSheetArr[indexDataList] = columns[indexDataList].length;
            if (columnHeadMerge) {
              dataJsonSheetArr[indexDataList] = itemDataList;
            } else {
              itemDataList.forEach(item => {
                let obj = {};
                columns[indexDataList].forEach(i => {
                  obj[i.name] = item[i.code];
                });
                dataJsonSheetArr[indexDataList].push(obj);
              });
            }
            resolve();
          });
        });
        Promise.all(promises).then(() => {
          // 配置文件类型
          const wopts = { bookType: 'xlsx', bookSST: true, type: 'binary', cellStyles: true };
          downloadExl(dataJsonSheetArr, wopts, option, columnsLenSheetArr, columnHeadMerge, columnHeader, columnHeaderGroup, columns, sheetName, fileName);
        });
      }

      // 最后一行列头映射
      function changeTitle(value, columnHeadMerge, columns) {
        if (!columnHeadMerge) {
          return value;
        } else {
          let title = value;
          columns.forEach(item => {
            if (item.code === value) {
              title = item.name;
            }
          });
          return title;
        }
      }
    
      // 导出方法配置项
      function downloadExl(json, type, option, columnsLen, columnHeadMerge, columnHeader, columnHeaderGroup, columns, sheetName, fileName) {
        const keyMapSheetArr = [];
        const tmpdatasSheetArr = [];
        const tmpdataSheetArr = [];
        const outputPosSheetArr = [];
        const promises = json.map((itemDataList, indexDataList) => {
          return new Promise((resolve) => {
            // 动态创建多个变量
            keyMapSheetArr[indexDataList] = []; // 获取keys
            tmpdatasSheetArr[indexDataList] = itemDataList[0];
            const borderAll = {
              //单元格外侧框线
              top: {
                style: 'thin',
              },
              bottom: {
                style: 'thin',
              },
              left: {
                style: 'thin',
              },
              right: {
                style: 'thin',
              },
            };
            const columnHeaderGroupLen = columnHeaderGroup.length;
            if (columnHeaderGroup.length) {
              if (columnHeadMerge) {
                itemDataList.unshift({}); // 向表格数据中插入1行位置(标题)
                for (let i = 0; i < columnHeaderGroup.length; i++) {
                  itemDataList.unshift({}); // 向表格数据中插入columnHeaderGroup.length行位置(第一行至倒数第二行列头)
                }
                for (const k in tmpdatasSheetArr[indexDataList]) {
                  // 为插入的X行位置添加数据
                  keyMapSheetArr[indexDataList].push(k);
                  columnHeaderGroup.forEach((item,index) => {
                    itemDataList[index][k] = item[k]; // 用于展示插入列头
                  });
                  itemDataList[columnHeaderGroupLen][k] = k; // 用于展示正常列头
                }
              } else {
                itemDataList.unshift({}); // 向表格数据中插入1行位置(标题)
                for (const k in tmpdatasSheetArr[indexDataList]) {
                  // 为插入的X行位置添加数据
                  keyMapSheetArr[indexDataList].push(k);
                  itemDataList[0][k] = k; // 插入第一行列头
                }
              }
            } else {
              if (columnHeadMerge) {
                if (Object.keys(columnHeader).length) {
                  itemDataList.unshift({}, {}); // 向表格数据中插入2行位置(标题和第一行列头)
                  for (const k in tmpdatasSheetArr[indexDataList]) {
                    // 为插入的X行位置添加数据
                    keyMapSheetArr[indexDataList].push(k);
                    itemDataList[0][k] = columnHeader[k]; // 用于展示插入列头
                    itemDataList[1][k] = k; // 用于展示正常列头
                  }
                } else {
                  itemDataList.unshift({}); // 向表格数据中插入1行位置(标题)
                  for (const k in tmpdatasSheetArr[indexDataList]) {
                    // 为插入的X行位置添加数据
                    keyMapSheetArr[indexDataList].push(k);
                    itemDataList[0][k] = k; // 插入第一行列头
                  }
                }
              } else {
                itemDataList.unshift({}); // 向表格数据中插入1行位置(标题)
                for (const k in tmpdatasSheetArr[indexDataList]) {
                  // 为插入的X行位置添加数据
                  keyMapSheetArr[indexDataList].push(k);
                  itemDataList[0][k] = k; // 插入第一行列头
                }
              }
            }
            console.log(itemDataList)
            console.log(columnsLen)
            console.log(columns)
            console.log(keyMapSheetArr)
            console.log(tmpdatasSheetArr)
            tmpdataSheetArr[indexDataList] = []; // 用来保存转换好的json
            itemDataList
              .map((v, i) => {
                const data = keyMapSheetArr[indexDataList].map((k, j) => {
                  return Object.assign(
                    {},
                    {
                      v: v[k],
                      position: (j > 25 ? getCharCol(j) : String.fromCharCode(65 + j)) + (i + 2), // 表格数据的位置
                    }
                  );
                });
                return data;
              })
              .reduce((prev, next) => prev.concat(next))
              .forEach(
                (v, i) =>
                  (tmpdataSheetArr[indexDataList][v.position] = {
                    v: changeTitle(v.v, columnHeadMerge, columns[indexDataList]),
                    s: {
                      font: { sz: option[indexDataList].fontSizeList },
                      alignment: { vertical: 'center', horizontal: 'center', wrapText: 'true' },
                      border: borderAll,
                    },
                  })
              );
            outputPosSheetArr[indexDataList] = Object.keys(tmpdataSheetArr[indexDataList]) // 设置区域,比如表格从A1到D10
            tmpdataSheetArr[indexDataList].A1 = { v: option[indexDataList].title }; // A1-A4区域的内容
            outputPosSheetArr[indexDataList] = ['A1'].concat(outputPosSheetArr[indexDataList]);
            // 对所有单元格列的样式批量处理
            Object.keys(tmpdataSheetArr[indexDataList]).forEach(item => {
              const lastChar = item.charAt(item.length - 1); // 获取坐标的最后一位字符(例如 A12 为 2)
              const isChar = isNaN(parseFloat(item.charAt(item.length - 2))); // 判断坐标的倒数第二位字符是否是非数字(例如 A2 为 true)
              if (columnHeadMerge) {
                // 判断列头和AA1等超出26列的情况
                // 设置从第2行到columnHeaderGroupLen+2行的表头样式
                const columnLen = columnHeaderGroupLen || 1;
                if ((parseFloat(lastChar) <= columnLen+2 && isChar)) {
                  tmpdataSheetArr[indexDataList][item].s = {
                    font: { sz: option[indexDataList].fontSize, bold: option[indexDataList].fontBold },
                    border: borderAll,
                    fill: { fgColor: { rgb: option[indexDataList].fillColor } },
                    alignment: {
                      vertical: option[indexDataList].alignmentVertical,
                      horizontal: option[indexDataList].alignmentHorizontal,
                      wrapText: true,
                    },
                  };
                }
              } else if (
                lastChar === `2` && isChar
              ) {
                tmpdataSheetArr[indexDataList][item].s = {
                  font: { sz: option[indexDataList].fontSize, bold: option[indexDataList].fontBold },
                  border: borderAll,
                  fill: { fgColor: { rgb: option[indexDataList].fillColor } },
                  alignment: { vertical: option[indexDataList].alignmentVertical, horizontal: option[indexDataList].alignmentHorizontal, wrapText: true },
                };
              }
            });
            // ======================在此处对某一列单元格样式进行单独==============================
            tmpdataSheetArr[indexDataList].A1.s = {
              font: { sz: option[indexDataList].fontSizeTitle, bold: true },
              border: borderAll,
              alignment: { vertical: 'center', horizontal: 'center' },
            };
        
            // 列级别样式修改
            if (option[indexDataList].styleRow && option[indexDataList].styleRow.length) {
              option[indexDataList].styleRow.forEach(item => {
                item.style.border = borderAll;
                item.cells.forEach(i => {
                  Object.keys(tmpdataSheetArr[indexDataList]).forEach(j => {
                    let num = /(\d+(\.\d+)?)/.exec(j);
                    if (num && num.length > 0) {
                      num = num[0];
                    }
                    if (i === num) {
                      tmpdataSheetArr[indexDataList][j].s = item.style;
                    }
                  });
                });
              });
            }
        
            // 单元格批量样式修改
            if (option[indexDataList].styleGroup && option[indexDataList].styleGroup.length) {
              option[indexDataList].styleGroup.forEach(item => {
                item.style.border = borderAll;
                item.cells.forEach(i => {
                  tmpdataSheetArr[indexDataList][i].s = item.style;
                });
              });
            }
        
            // s-e 代表区域 c-r 代表列-行的索引
            // 定制化改动地方
            let mergesLen = columnsLen[indexDataList] - 1;
            tmpdataSheetArr[indexDataList]['!merges'] = [
              {
                s: { c: 0, r: 0 },
                e: { c: mergesLen, r: 0 },
              },
            ]; // <====合并单元格
        
            // 设置单元格合并
            if (option[indexDataList].merges && option[indexDataList].merges.length) {
              tmpdataSheetArr[indexDataList]['!merges'] = tmpdataSheetArr[indexDataList]['!merges'].concat(option[indexDataList].merges);
            }
        
            let dataArrWidth = [];
            for (let i = 0; i < columnsLen[indexDataList] + 1; i++) {
              dataArrWidth.push({ wpx: option[indexDataList].width || 150 });
            }
            // 例 dataArrWidth[3].wpx = 130;
            // 以列为维度，设置单元格宽度
            if (option[indexDataList].widthColums && option[indexDataList].widthColums.length) {
              option[indexDataList].widthColums.forEach(item => {
                item.cells.forEach(i => {
                  dataArrWidth[i].wpx = item.width;
                });
              });
            }
            tmpdataSheetArr[indexDataList]['!cols'] = dataArrWidth; // <====设置一列宽度， 代表20列都是300宽
            resolve();
          });
        });
        Promise.all(promises).then(() => {
          const tmpWB = {
            SheetNames: sheetName, // sheet名
            Sheets: {},
          };
          console.log(tmpdataSheetArr)
          console.log(outputPosSheetArr)
          tmpdataSheetArr.forEach((item, index) => {
            tmpWB.Sheets[sheetName[index]] = Object.assign(
              {},
              tmpdataSheetArr[index], // 内容
              {
                '!ref': `${outputPosSheetArr[index][0]}:${outputPosSheetArr[index][outputPosSheetArr[index].length - 1]}`, // 设置填充区域(表格渲染区域)
              }
            )
          });
          const tmpDown = new Blob(
            [
              s2ab(
                XLSX.write(
                  tmpWB,
                  { bookType: type == undefined ? 'xlsx' : type.bookType, bookSST: false, type: 'binary' } // 这里的数据是用来定义导出的格式类型
                )
              ),
            ],
            {
              type: '',
            }
          );
          // 定制化改动地方
          saveAs(tmpDown, `${fileName}.${type.bookType == 'biff2' ? 'xls' : type.bookType}`);
        });
      }
    
      // 导出IE兼容
      function saveAs(obj, fileName) {
        const tmpa = document.createElement('a');
        tmpa.download = fileName || '未命名';
        // 兼容ie
        if ('msSaveOrOpenBlob' in navigator) {
          window.navigator.msSaveOrOpenBlob(obj, 'excle文件名' + '.xlsx');
        } else {
          tmpa.href = URL.createObjectURL(obj);
        }
        tmpa.click();
        setTimeout(function() {
          URL.revokeObjectURL(obj);
        }, 100);
      }
    
      function s2ab(s) {
        if (typeof ArrayBuffer !== 'undefined') {
          const buf = new ArrayBuffer(s.length);
          const view = new Uint8Array(buf);
          for (let i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
          return buf;
        } else {
          const buf = new Array(s.length);
          for (let i = 0; i != s.length; ++i) buf[i] = s.charCodeAt(i) & 0xff;
          return buf;
        }
      }
    
      // 获取26个英文字母用来表示excel的列
      function getCharCol(n) {
        const temCol = '';
        let s = '';
        let m = 0;
        while (n > 0) {
          m = (n % 26) + 1;
          s = String.fromCharCode(m + 64) + s;
          n = (n - m) / 26;
        }
        return s;
      }
      handleExportData();
    };
    
    export default ExportExcel;
    