// 엑셀파일 다중 생성
// 'use strict'

// const XLSX = require('xlsx');

// let finalHeaders = ['colA', 'colB', 'colC'];
// let data = [
//     [ { colA: 1, colB: 2, colC: 3 }, { colA: 4, colB: 5, colC: 6 }, { colA: 7, colB: 8, colC: 9 } ],
//     [ { colA:11, colB:12, colC:13 }, { colA:14, colB:15, colC:16 }, { colA:17, colB:18, colC:19 } ],
//     [ { colA:21, colB:22, colC:23 }, { colA:24, colB:25, colC:26 }, { colA:27, colB:28, colC:29 } ]
// ];
// console.log(data);
// data.forEach((array, i) => {
//     let ws = XLSX.utils.json_to_sheet(array, {header: finalHeaders});
//     let wb = XLSX.utils.book_new()
//     XLSX.utils.book_append_sheet(wb, ws, "SheetJS")
//     let exportFileName = `workbook_${i}.xls`;
//     XLSX.writeFile(wb, exportFileName)
// });

//엑셀파일 단일 생성
'use strict'

const XLSX = require('xlsx');

let finalHeaders = ['colA', 'colB', 'colC'];
let data = [ { colA: 1, colB: 2, colC: 3 }, { colA: 4, colB: 5, colC: 6 }, { colA: 7, colB: 8, colC: 9 } ]

console.log(data);
    let ws = XLSX.utils.json_to_sheet(data, {header: finalHeaders});
    let wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, "SheetJS")
    let exportFileName = `workbook.xls`;
    XLSX.writeFile(wb, exportFileName)
