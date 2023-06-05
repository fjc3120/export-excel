/*
 * @Descripttion: 前端导出组件
 * @Author: 小鸟游露露
 * @Date: 2021-06-02 10:28:40
 * @LastEditTime: 2023-04-17 10:15:13
 * @Copyright: Copyright (c) 2018, Hand
 * 实参obj为一个对象，包含以下元素: Array-dataList(表格数据)  Object-option(配置信息) Array-columnsList(列头信息) Boolean-columnHeadMerge(是否开启列头合并功能) Array-columnsList(最后一行表列头映射表)  Array-columnHeader(第一行表头信息) Array-columnHeaderGroup(第一行至倒数第二行表头信息)

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
            title = '',
            columnHeadMerge = false,
            columnHeader = {},
            columnHeaderGroup = [],
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
          if (title) {
            option.title = title;
          }
          let colList = [];
          columnsList.forEach(item => {
            colList.push(item.code);
          });
          let dataAll = JSON.parse(JSON.stringify(dataList));
          // 根据columnsList补全dataList缺失的字段
          for (let i = 0; i < dataAll.length; i++) {
            let row = dataAll[i];
            for (let j = 0; j < colList.length; j++) {
              let key = colList[j];
              if (!Object.prototype.hasOwnProperty.call(row, key)) {
                row[key] = '';
              }
            }
          }
          // 对dataList每行数据的key进行排序(不重写会导致列的顺序错乱)
          let data = [];
          dataAll.forEach(item => {
            const objItem = {};
            colList.forEach(i => {
              objItem[i] = item[i];
            });
            data.push(objItem);
          });
          let dataExport = data.map((item) => {
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
          exportRun(dataExport, option, columnsList, columnHeadMerge, columnHeader, columnHeaderGroup); //数据-配置信息-最后一行列头信息-是否开启列头合并-列头第一行-列头第一行至倒数第二行
        } else {
          if (!("Notification" in window)) {
            alert("无导出数据");
          }
          else if (Notification.permission === "granted") {
            var notification = new Notification("无导出数据");
          }
          else if (Notification.permission !== 'denied') {
            Notification.requestPermission(function (permission) {
              if (permission === "granted") {
                var notification = new Notification("无导出数据");
              }
            });
          }
        }
      }
    
      // 执行导出方法
      function exportRun(data, option, columns, columnHeadMerge, columnHeader, columnHeaderGroup) {
        if (data.length) {
          let dataJson = [];
          let columnsLen = columns.length;
          if (columnHeadMerge) {
            dataJson = data;
          } else {
            data.forEach(item => {
              let obj = {};
              columns.forEach(i => {
                obj[i.name] = item[i.code];
              });
              dataJson.push(obj);
            });
          }
          // 配置文件类型
          const wopts = { bookType: 'xlsx', bookSST: true, type: 'binary', cellStyles: true };
          downloadExl(dataJson, wopts, option, columnsLen, columnHeadMerge, columnHeader, columnHeaderGroup, columns);
        } else {
          if (!("Notification" in window)) {
            alert("无导出数据");
          }
          else if (Notification.permission === "granted") {
            var notification = new Notification("无导出数据");
          }
          else if (Notification.permission !== 'denied') {
            Notification.requestPermission(function (permission) {
              if (permission === "granted") {
                var notification = new Notification("无导出数据");
              }
            });
          }
        }
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
      function downloadExl(json, type, option, columnsLen, columnHeadMerge, columnHeader, columnHeaderGroup, columns) {
        const borderAll = {
          // 单元格外侧框线
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
        const keyMap = []; // 获取keys
        let tmpdatas = json[0];
        const columnHeaderGroupLen = columnHeaderGroup.length;
        if (columnHeaderGroup.length) {
          if (columnHeadMerge) {
            json.unshift({}); // 向表格数据中插入1行位置(标题)
            for (let i = 0; i < columnHeaderGroup.length; i++) {
              json.unshift({}); // 向表格数据中插入columnHeaderGroup.length行位置(第一行至倒数第二行列头)
            }
            for (const k in tmpdatas) {
              // 为插入的X行位置添加数据
              keyMap.push(k);
              columnHeaderGroup.forEach((item,index) => {
                json[index][k] = item[k]; // 用于展示插入列头
              });
              json[columnHeaderGroupLen][k] = k; // 用于展示正常列头
            }
          } else {
            json.unshift({}); // 向表格数据中插入1行位置(标题)
            for (const k in tmpdatas) {
              // 为插入的X行位置添加数据
              keyMap.push(k);
              json[0][k] = k; // 插入第一行列头
            }
          }
        } else {
          if (columnHeadMerge) {
            if (Object.keys(columnHeader).length) {
              json.unshift({}, {}); // 向表格数据中插入2行位置(标题和第一行列头)
              for (const k in tmpdatas) {
                // 为插入的X行位置添加数据
                keyMap.push(k);
                json[0][k] = columnHeader[k]; // 用于展示插入列头
                json[1][k] = k; // 用于展示正常列头
              }
            } else {
              json.unshift({}); // 向表格数据中插入1行位置(标题)
              for (const k in tmpdatas) {
                // 为插入的X行位置添加数据
                keyMap.push(k);
                json[0][k] = k; // 插入第一行列头
              }
            }
          } else {
            json.unshift({}); // 向表格数据中插入1行位置(标题)
            for (const k in tmpdatas) {
              // 为插入的X行位置添加数据
              keyMap.push(k);
              json[0][k] = k; // 插入第一行列头
            }
          }
        }
        let tmpdata = []; // 用来保存转换好的json
        json
          .map((v, i) => {
            const data = keyMap.map((k, j) => {
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
              (tmpdata[v.position] = {
                v: changeTitle(v.v, columnHeadMerge, columns),
                s: {
                  font: { sz: option.fontSizeList },
                  alignment: { vertical: 'center', horizontal: 'center', wrapText: 'true' },
                  border: borderAll,
                },
              })
          );
        let outputPos = Object.keys(tmpdata); // 设置区域,比如表格从A1到D10
        tmpdata.A1 = { v: option.title }; // A1-A4区域的内容
        outputPos = ['A1'].concat(outputPos);
        // 对所有单元格列的样式批量处理
        Object.keys(tmpdata).forEach(item => {
          const lastChar = item.charAt(item.length - 1); // 获取坐标的最后一位字符(例如 A12 为 2)
          const isChar = isNaN(parseFloat(item.charAt(item.length - 2))); // 判断坐标的倒数第二位字符是否是非数字(例如 A2 为 true)
          if (columnHeadMerge) {
            // 判断列头和AA1等超出26列的情况
            // 设置从第2行到columnLen+2行的表头样式
            const columnLen = columnHeaderGroupLen || 1;
            if ((parseFloat(lastChar) <= columnLen+2 && isChar)) {
              tmpdata[item].s = {
                font: { sz: option.fontSize, bold: option.fontBold },
                border: borderAll,
                fill: { fgColor: { rgb: option.fillColor } },
                alignment: {
                  vertical: option.alignmentVertical,
                  horizontal: option.alignmentHorizontal,
                  wrapText: true,
                },
              };
            }
          } else if (
            lastChar === `2` && isChar
          ) {
            tmpdata[item].s = {
              font: { sz: option.fontSize, bold: option.fontBold },
              border: borderAll,
              fill: { fgColor: { rgb: option.fillColor } },
              alignment: { vertical: option.alignmentVertical, horizontal: option.alignmentHorizontal, wrapText: true },
            };
          }
        });
        // ======================在此处对某一列单元格样式进行单独处理==============================
        tmpdata.A1.s = {
          font: { sz: option.fontSizeTitle, bold: true },
          border: borderAll,
          alignment: { vertical: 'center', horizontal: 'center' },
        };
        // 列级别样式修改
        if (option.styleRow && option.styleRow.length) {
          option.styleRow.forEach(item => {
            item.style.border = borderAll;
            item.cells.forEach(i => {
              Object.keys(tmpdata).forEach(j => {
                let num = /(\d+(\.\d+)?)/.exec(j);
                if (num && num.length > 0) {
                    num = num[0];
                }
                if (i === num) {
                  tmpdata[j].s = item.style;
                }
              });
            });
          });
        }

        // 单元格批量样式修改
        if (option.styleGroup && option.styleGroup.length) {
          option.styleGroup.forEach(item => {
            item.style.border = borderAll;
            item.cells.forEach(i => {
              tmpdata[i].s = item.style;
            });
          });
        }
        // s-e 代表区域 c-r 代表列-行的索引
        let mergesLen = columnsLen - 1;
        tmpdata['!merges'] = [
          {
            s: { c: 0, r: 0 },
            e: { c: mergesLen, r: 0 },
          },
        ]; // <====合并单元格

        // 设置单元格合并
        if (option.merges && option.merges.length) {
          tmpdata['!merges'] = tmpdata['!merges'].concat(option.merges);
        }

        let dataArrWidth = [];
        for (let i = 0; i < columnsLen + 1; i++) {
          dataArrWidth.push({ wpx: option.width || 150 });
        }

        // 例 dataArrWidth[3].wpx = 130;
        // 以列为维度，设置单元格宽度
        if (option.widthColums && option.widthColums.length) {
          option.widthColums.forEach(item => {
            item.cells.forEach(i => {
              dataArrWidth[i].wpx = item.width;
            });
          });
        }
        tmpdata['!cols'] = dataArrWidth; // <====设置一列宽度， 代表20列都是300宽
        const tmpWB = {
          SheetNames: ['mySheet'], // 保存的表标题
          Sheets: {
            mySheet: Object.assign(
              {},
              tmpdata, // 内容
              {
                '!ref': `${outputPos[0]}:${outputPos[outputPos.length - 1]}`, // 设置填充区域(表格渲染区域)
              }
            ),
          },
        };
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
        saveAs(tmpDown, `${option.title}.${type.bookType == 'biff2' ? 'xls' : type.bookType}`);
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
    