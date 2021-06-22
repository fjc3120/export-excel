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
    var XLSX = _interopRequireDefault(require("./xlsx-style"));
    function ExportExcel(obj={}) {
      // 点击导出按钮
      function handleExportData() {
        if (Object.keys(obj).length) {
          let { dataList = [], option = {}, columnsList = [] } = obj;
          exportRun(dataList, option, columnsList); // 数据-配置信息-列头信息
        }
      }
    
      // 执行导出方法
      function exportRun(data, option, columns) {
        if (data.length) {
          let dataJson = [];
          let len = columns.length;
          data.forEach(item => {
            let obj = {};
            columns.forEach(i => {
              obj[i.name] = item[i.code];
            });
            dataJson.push(obj);
          });
          // 配置文件类型
          const wopts = { bookType: 'xlsx', bookSST: true, type: 'binary', cellStyles: true };
          downloadExl(dataJson, wopts, option, len);
        }
      }
    
      // 导出方法配置项
      function downloadExl(json, type, option, columnsLen) {
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
        let tmpdatas = json[0];
        json.unshift({}); // 向表格数据中插入4行位置(标题和参数)
        const keyMap = []; // 获取keys
        for (const k in tmpdatas) {
          // 为插入的5行位置添加数据
          keyMap.push(k);
          json[0][k] = k; // json[3][k] = k 3为插入5行的最后一行索引，用于展示列头
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
                v: v.v,
                s: {
                  font: { sz: 10 },
                  alignment: { vertical: 'center', horizontal: 'center', wrapText: 'true' },
                  border: borderAll,
                },
              })
          );
        let outputPos = Object.keys(tmpdata); // 设置区域,比如表格从A1到D10
        tmpdata.A1 = { v: option.title }; // A1-A4区域的内容
        outputPos = ['A1'].concat(outputPos);
        // 对所有单元格列的样式批量处理
        console.log(tmpdata)
        Object.keys(tmpdata).forEach(item => {
          if (item.charAt(item.length - 1) === '2' && isNaN(parseFloat(item.charAt(item.length - 2)))) {
            tmpdata[item].s = {
              font: { sz: option.fontSize, bold: option.fontBold },
              border: borderAll,
              fill: { fgColor: { rgb: option.fillColor } },
              alignment: { vertical: option.alignmentVertical, horizontal: option.alignmentHorizontal },
            };
          }
        });
        // ======================在此处对某一列单元格样式进行单独处理==============================
        tmpdata.A1.s = {
          font: { sz: 14, bold: true },
          border: borderAll,
          alignment: { vertical: 'center', horizontal: 'center' },
        };
        // s-e 代表区域 c-r 代表列-行的索引
        let mergesLen = columnsLen - 1;
        tmpdata['!merges'] = [
          {
            s: { c: 0, r: 0 },
            e: { c: mergesLen, r: 0 },
          },
        ]; // <====合并单元格
        let dataArrWidth = [];
        for (let i = 0; i < columnsLen + 1; i++) {
          dataArrWidth.push({ wpx: option.width });
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
    