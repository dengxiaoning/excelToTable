<template>
  <!-- <div v-html="rowsHtml"
       class="continertable"></div> -->
  <el-table v-if="resTableData.length"
            :data="resTableData"
            size='small'
            height="950"
            :span-method="objectSpanMethod">
    <el-table-column v-for="(item,index) in tableHeadArr"
                     :key="index"
                     :label="item.label"
                     align="center">
      <el-table-column v-for="(sItem,sIndex) in item.children"
                       :key="index+'_'+sIndex"
                       :prop="sItem.prop"
                       :label="sItem.label"
                       show-overflow-tooltip
                       fixed="left"
                       :width="sIndex<2?80:150">
      </el-table-column>
      <el-table-column v-if="headerGroup"
                       :label="headerGroup.label"
                       :prop="headerGroup.prop"
                       align="center">
        <el-table-column v-for="(hItem,hIndex) in headerGroup.children"
                         :key="index+'_090_'+hIndex"
                         :label="hItem.label"
                         :prop="hItem.prop"
                         width="150">
        </el-table-column>
      </el-table-column>
      <el-table-column v-for="(tItem,tIndex) in item.children.children"
                       :key="index+'_01_'+tIndex"
                       :label="tItem.label"
                       :prop="tItem.prop"
                       align="center">
        <el-table-column v-for="(fItem,fIndex) in tItem.children"
                         :key="index+'_'+tIndex+'_'+fIndex"
                         :label="fItem.label"
                         :prop="fItem.prop"
                         width="150">
        </el-table-column>
      </el-table-column>
      <el-table-column v-if="remackArr"
                       :prop="remackArr.prop"
                       :label="remackArr.label"
                       show-overflow-tooltip
                       width="150">
      </el-table-column>
    </el-table-column>
  </el-table>
</template>

<script>
import { read, utils, writeFileXLSX } from 'xlsx';
var wtregex = /(^\s|\s$|\n)/;
var decregex = /[&<>'"]/g;
var htmlcharegex = /[\u0000-\u001f]/g;
const propRegex = /[\w]{3}-[\w]{1}/;
const rowNumRegex = /\d{1,}$/;
const colNumArr = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
export default {
  data () {
    return {
      rowsHtml: '',
      globalType: 'base',
      sheetTabs: 'base',
      cacheWB: {},
      categoryCN: {},
      existMergeRow: {},
      existMergeCol: {},
      mergeColAtRow: {},
      mergeKeys: [],
      tableHeadArr: [],
      resTableData: [],
      remackArr: undefined,
      headerGroup: undefined
    }
  },
  created () {
    this.loadExcel();
  },
  watch: {
    sheetTabs (val) {
      if (val === 'base') {
        this.remackArr = undefined
        this.headerGroup = undefined
      } else if (val === 'fun') {
        this.remackArr = this.tableHeadArr[0]['children'][6]
        this.tableHeadArr[0]['children'][6] = this.tableHeadArr[0]['children'][7]
        this.tableHeadArr[0]['children'].pop();
        // 然后把五个元素变成header group，tableHeadArr先前移动一位
        this.headerGroup = this.tableHeadArr[0]['children'][5];
        this.tableHeadArr[0]['children'][5] = this.tableHeadArr[0]['children'][6]
        this.tableHeadArr[0]['children'].pop();
        this.headerGroup['children'] = [this.tableHeadArr[0]['children']['children'][0]['children'].shift()]
      }
    }
  },
  methods: {
    objectSpanMethod ({ rowIndex, columnIndex }) {
      // 需要合并的列
      const mergeRow = this.categoryCN[columnIndex];
      // 合并列先找到那一例属于哪一行
      const mergeColAtRow = this.mergeColAtRow[rowIndex];
      // 通过列index 获取这一列，哪些行有合并行为
      const existMergeRow = this.existMergeRow[columnIndex]
      const existMergeCol = this.existMergeCol[columnIndex]
      if (mergeRow) {
        let rowspanObj = mergeRow.find((ee, index) => (ee.rowNum - 4) == rowIndex)
        // 第一列
        if (columnIndex == 0) {

          // 判断改行是否参与合并行为
          if (existMergeRow[rowIndex]) {
            if (rowspanObj) {
              const _row = rowspanObj.rowspan
              const _col = rowspanObj.colspan
              return {
                rowspan: _row ? _row : _col ? 1 : 0,
                colspan: _col ? _col : _row ? 1 : 0
              }
            } else {
              return {
                rowspan: 0,
                colspan: 0
              }
            }
          }
        } else {  // 其它列,合并操作
          // 有行合并也有列合并
          if (existMergeRow && existMergeRow[rowIndex] && existMergeCol && existMergeCol[columnIndex]) {
            if (rowspanObj) {
              const _row = rowspanObj.rowspan
              const _col = rowspanObj.colspan
              return {
                rowspan: _row ? _row : _col ? 1 : 0,
                colspan: _col ? _col : _row ? 1 : 0
              }
            } else {
              return {
                rowspan: 0,
                colspan: 0
              }
            }
          } else if (existMergeRow && existMergeRow[rowIndex]) {// 只有行合并行为

            if (rowspanObj) {
              const _row = rowspanObj.rowspan
              const _col = rowspanObj.colspan
              return {
                rowspan: _row ? _row : _col ? 1 : 0,
                colspan: _col ? _col : _row ? 1 : 0
              }

            } else {
              // 除开第一列，其它行进行合并后，未合并的行需要隐藏，避免出现多行错乱问题
              return {
                rowspan: 0,
                colspan: 0
              }
            }
          }
        }
      }

      // 找合并列对应的那一行，避免所有列做重复操作
      if (mergeColAtRow && mergeColAtRow[columnIndex]) {
        const mergeRowPlus = this.categoryCN[columnIndex];
        if (mergeRowPlus) {
          // 只有列表合并，
          let colspanObj = mergeRowPlus.find((ee, index) => ee.colNum == columnIndex && (ee.rowNum - 4) == rowIndex)
          if (colspanObj) {
            const _col = colspanObj.colspan
            return {
              rowspan: 1,
              colspan: _col
            }
          }
        } else {
          return {
            rowspan: 0,
            colspan: 0
          }
        }
      }
    },
    async loadExcel () {

      try {

        /* Download from https://docs.sheetjs.com/pres.numbers */
        const f = await fetch('小程序开发任务清单.xlsx');
        const ab = await f.arrayBuffer();
        /* parse workbook */
        const wb = read(ab);
        console.log(wb);
        this.cacheWB['base'] = wb.Sheets[wb.SheetNames[0]];
        this.cacheWB['fun'] = wb.Sheets[wb.SheetNames[1]];
        /* update data */
        let htmlStr = utils.sheet_to_html(wb.Sheets[wb.SheetNames[0]], {
          id: "dynamictable"
        })
        htmlStr = htmlStr.replace('<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body>', '').replace('</body></html>', '')

        this.rowsHtml = htmlStr;
        this.dealTableData(wb.Sheets[wb.SheetNames[0]])
      } catch (error) {
        console.log(error);
      }
    },
    dealTableData (worksheet) {
      // const worksheet = workbook.Sheets[sheetNames[0]];
      // 拿到这张表中表格数据的范围，
      const range = utils.decode_range(worksheet['!ref']);
      // console.log(worksheet['!ref']);  // A1:E5
      //保存数据范围数据
      const row_start = range.s.r; // 表格范围，开始行的数据
      const row_end = range.e.r; // 表格范围，结束行的数据
      const col_start = range.s.c; // 表格范围，开始列的数据
      const col_end = range.e.c; // 表格范围，结束行的数据
      const tableMerge = worksheet['!merges'] || []; // 表格中进行单元格合并操作的数据
      var oo = [];
      var tableArr = []; // 存储所以的td 数组
      var preamble = "<tr>"; // 转 html 时进行拼接
      // let rows = [], row_data, i, addr, cell;
      //按行对 sheet 内的数据循环
      //首先读取当前对象内的所有行数据，从开始到结束
      for (var R = row_start; R <= row_end; ++R) {
        var innerRow = []
        var innerRowJson = []
        // out.push(make_html_row(ws, r, R, o));
        // 读取列数据，开始到结束
        for (var C = col_start; C <= col_end; ++C) {
          var RS = 0, CS = 0;
          // 针对表中进行合并单元格操作的数据
          for (var j = 0; j < tableMerge.length; ++j) {
            if (tableMerge[j].s.r > R || tableMerge[j].s.c > C) continue;
            if (tableMerge[j].e.r < R || tableMerge[j].e.c < C) continue;
            if (tableMerge[j].s.r < R || tableMerge[j].s.c < C) { RS = -1; break; }
            RS = tableMerge[j].e.r - tableMerge[j].s.r + 1; CS = tableMerge[j].e.c - tableMerge[j].s.c + 1; break;
          }
          if (RS < 0) continue;
          var coord = utils.encode_cell({ r: R, c: C });
          var cell = worksheet[coord];
          // console.log(cell);
          var sp = ({});
          if (RS > 1) sp.rowspan = RS;
          if (CS > 1) sp.colspan = CS;
          sp.t = cell && cell.t || 'z';
          sp.id = "sjs" + "-" + coord;
          if (sp.t != "z") { sp.v = cell.v; if (cell.z != null) sp.z = cell.z; }
          // 这里就得到了我们所要的数据
          var w = (cell && cell.v != null) && (cell.h || this.escapehtml(cell.w || (utils.format_cell(cell), cell.w) || "")) || "";

          innerRow.push(this.writextag('td', w, sp));
          // 这是一行数据，中的多个td，要将其转换为elementui 的数据格式
          if (w) { // 排除空数据
            innerRowJson.push({ name: w, ...sp })
          }

        }
        let str = preamble + innerRow.join("") + "</tr>";
        oo.push(str)
        if (innerRowJson.length > 0) { tableArr.push(innerRowJson) }

      }

      // 组装表头
      this.assemblyTableData(tableArr);

      this.asseblyTableColumn(tableArr);
    },
    // 组装一个表单类字段
    asseblyTableColumn (arr) {
      const firstArr = arr[0];
      const secondArr = arr[1];
      const thirdArr = arr[2];
      const secondChildren = []
      let thirdObj = {}
      const thirdChildren = []
      // 数组第一个为表头
      const resArr = [{
        label: firstArr[0]['name'],
        ...firstArr[0]
      }];

      thirdArr.forEach(e => {
        let propStr = e.id.match(propRegex)
        thirdChildren.push({
          label: e.name,
          prop: propStr[0],
          ...e
        })
      })
      secondArr.forEach(e => {
        let propStr = e.id.match(propRegex)
        // 包含colspan为下一个表头
        if (e.hasOwnProperty('colspan')) {
          thirdObj = {
            label: this.removeHTMLTags(e.name),
            ...e,
            children: thirdChildren,
          }
        } else {

          secondChildren.push({
            label: e.name,
            prop: propStr[0],
            ...e
          })
        }
      })
      secondChildren['children'] = [thirdObj];
      resArr[0]['children'] = secondChildren;
      this.tableHeadArr = resArr
      // 设置类型
      this.sheetTabs = this.globalType;
    },
    /**
     * 组装表单数据
     * @description
     * 由于数据中存在行/列合并的情况
     * 数组的长度完全不同，而elementUI 加载数据
     * 是通过prop 进行一一绑定的，
     * 所以我们就必须找到每一个原始中的唯一标识来进行绑定
     * 经过对数据的仔细分析发现，xlsx使用了表格字母作为id
     * 于是我们可以通过这个来进行数据绑定
     */
    assemblyTableData (arr) {
      this.categoryCN = [];
      /**
       * 合并信息
       * [{
       * rowNum:0,
       * colNum:0,
       * rowspan:0,
       * colspan:0,
       * }]
       */
      const spInfo = [];
      this.resTableData = [];
      // 从第三一个开始
      for (let i = 3; i < arr.length; i++) {
        let eachObj = {}
        let rowcolObj = {}
        arr[i].forEach(e => {
          let propStr = e.id.match(propRegex)
          eachObj[propStr[0]] = this.removeHTMLTags(e.name);
          const colName = propStr[0].match(/\w$/)[0];
          const rowNum = e.id.match(rowNumRegex)[0]

          if (e.hasOwnProperty('rowspan') && e.hasOwnProperty('colspan')) {
            spInfo.push({
              rowspan: e['rowspan'],
              colspan: e['colspan'],
              rowNum,
              colNum: colNumArr.findIndex(e => e === colName)
            })
          } else if (e.hasOwnProperty('rowspan')) {
            spInfo.push({
              rowspan: e['rowspan'],
              rowNum,
              colNum: colNumArr.findIndex(e => e === colName)
            })
          } else if (e.hasOwnProperty('colspan')) {
            spInfo.push({
              colspan: e['colspan'],
              rowNum,
              colNum: colNumArr.findIndex(e => e === colName)
            })
          }

        })
        this.resTableData.push(eachObj)
      }
      const categoryCN = this.categoryCN;
      // 根据colNum进行分组，便于合并
      spInfo.forEach(cn => {
        if (categoryCN[cn.colNum]) {
          categoryCN[cn.colNum].push(cn)
        } else {
          categoryCN[cn.colNum] = [cn];
        }
      })

      this.existMergeRow = {};
      this.existMergeCol = {};
      this.mergeColAtRow = {};
      const existMergeRow = this.existMergeRow// 合并行行为记录
      const existMergeCol = this.existMergeCol // 合并列行为记录
      const mergeColAtRow = this.mergeColAtRow;// 记录合并列是属于哪一行的
      // 计算出哪些行有合并参与合并行为，因为有的行完全没有合并操作，在table merge操作是需要特殊处理renter {rowspan:1,colspan:1}
      Object.keys(categoryCN).map(key => {
        categoryCN[key].forEach((e) => {
          // 有e.rowspan才执行如下操作
          if (e.rowspan) {
            if (!existMergeRow[key]) { existMergeRow[key] = {} }
            let rowspanNum = parseInt(e.rowspan);
            const rowNum = e.rowNum - 4;// 行数
            existMergeRow[key][rowNum] = true;// 标识他合并了
            let step = 1;
            while (rowspanNum > 1) {
              existMergeRow[key][rowNum + step] = true;// 标识他合并了
              step++;
              rowspanNum -= 1;
            }
          }

          if (e.colspan) {
            const keynum = parseInt(key) - 1
            const rowNumm = e.rowNum - 4;
            const colNum = e.colNum;// 列数
            // 行号:列号
            if (!mergeColAtRow[rowNumm]) { mergeColAtRow[rowNumm] = {} }
            mergeColAtRow[rowNumm][colNum] = true;


            if (!existMergeCol[keynum]) { existMergeCol[keynum] = {} }
            let colspanNum = parseInt(e.colspan);

            existMergeCol[keynum][colNum] = true;// 标识他合并了
            let colstep = 1;
            while (colspanNum > 1) {
              existMergeCol[keynum][colNum + colstep] = true;// 标识他合并了
              mergeColAtRow[rowNumm][colNum + colstep] = true;
              colstep++;
              colspanNum -= 1;
            }
          }
        })
      })

    },
    // js去除string里面html代码
    removeHTMLTags (str) {
      return str.replace(/<[^>]*>|(&#[\d|\w]{1,});/g, '');
    },
    coluMergeNum (existMergeRow, rowNum, rowspanNum) {
      function merge () {
        const rr = rowNum;
        let ww = rowspanNum;
        while (ww > -1) {
          existMergeRow[rr + 1] = true;// 标识他合并了
          --ww;
        }
      }
      merge()
    },
    evert (obj) {
      var o = ([]), K = keys(obj);
      for (var i = 0; i !== K.length; ++i) o[obj[K[i]]] = K[i];
      return o;
    },
    keys (o) {
      var ks = Object.keys(o), o2 = [];
      for (var i = 0; i < ks.length; ++i) {
        if (Object.prototype.hasOwnProperty.call(o, ks[i])) o2.push(ks[i]);
      }
      return o2;
    },

    rencoding () {
      var encodings = {
        '&quot;': '"',
        '&apos;': "'",
        '&gt;': '>',
        '&lt;': '<',
        '&amp;': '&'
      };
      return this.evert(encodings);
    },
    escapehtml (text) {
      var s = text + '';
      return s.replace(decregex, function (y) {
        return this.rencoding[y];
      }).replace(/\n/g, "<br/>")
        .replace(htmlcharegex, function (s) {
          return "&#x" + ("000" + s.charCodeAt(0).toString(16)).slice(-4) + ";";
        });
    },
    wxt_helper (h) {
      // console.log('h', h);
      return this.keys(h).map(function (k) {
        return " " + k + '="' + h[k] + '"';
      }).join("");
    },
    writextag (f, g, h) {
      return '<' + f + ((h != null) ? this.wxt_helper(h) : "")
        + ((g != null) ? (g.match(wtregex) ? ' xml:space="preserve"' : "")
          + '><p>' + g + '</p></' + f : "/") + '>';
    },
    switchTableData (type) {
      this.globalType = type;
      // 切换数据
      this.dealTableData(this.cacheWB[type])

      // this.rowsHtml = utils.sheet_to_html(this.cacheWB[type], {
      //   id: "dynamictable"
      // })
    },
    /* get live table and export to XLSX */
    exportFile () {
      const wb = utils.table_to_book(tableau.value.getElementsByTagName("TABLE")[0])
      writeFileXLSX(wb, "SheetJSVueHTML.xlsx");
    }
  }
}
</script>
<style lang="scss">
@import "./table.css";
#dynamictable {
  border-collapse: collapse; /* 使边框合并为单一边框 */
  td {
    height: 40px;
    line-height: 30px;
    min-width: 90px;
    max-width: 700px;
    border: 1px solid #215e97; /* 设置单元格边框 */
  }
  th {
    border: 1px solid #215e97; /* 设置单元格边框 */
  }
  tbody {
    tr {
      text-align: center;
      &:first-of-type {
        td {
          height: 50px;
          line-height: 50px;
          border: 1px solid #215e97;
        }
        background-color: #0b2b47;
      }
    }
  }
}
</style>