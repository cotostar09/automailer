//database conn
const mongojs = require('mongojs');
const db = mongojs('autoMailCheck', ['checkSendOk']);

const XLSX = require("xlsx");


db.checkSendOk.find( function(error,data){
    console.log(data);
    let finalHeaders = ['mail', 'fileName','check'];
    
    //엑셀 저장
    let ws = XLSX.utils.json_to_sheet(data, {header: finalHeaders});
    let wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SheetJS")
    let exportFileName = `./upload/sendDataOut.xlsx`;
    XLSX.writeFile(wb, exportFileName)
});