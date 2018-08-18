const nodemailer = require('nodemailer');
const smtpPool = require('nodemailer-smtp-pool');
const config = require('./config/config.json');
const fs = require('fs');
const XLSX = require("xlsx");

 
// let workbook = XLSX.readFile("./upload/note.xlsx")
// let worksheet = workbook.Sheets["Sheet1"]
// console.log(worksheet)

// let workbook = XLSX.readFile("./upload/note.xlsx")

// let sheet = workbook.SheetNames[0]; // 배열이므로 .length를 사용하여 갯수 확인가능
// let worksheet = workbook.Sheets[sheet];

// try{
// 	console.log( worksheet["A1"].v); //안의 데이터는 .v 를 사용해야 함
// }
// catch(exception){ //셀에 데이터가 없으면 exception이 남
// 	console.log("nono")
// }
// console.log(worksheet["!ref"]) // 사용한 셀의 길이



const workbook = XLSX.readFile("./upload/note.xlsx");
const sheetnames = Object.keys(workbook.Sheets);
 
let i = sheetnames.length;

const resData = {};
while (i--) {
    const sheetname = sheetnames[i];
    resData[sheetname] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetname]);
}
// console.log('resData');
// console.log(resData);

var keys = Object.keys(resData);    //시트 이름 추출
var sendData = resData[keys[0]];    //첫번째 시트 데이터

// console.log(keys);

var sendKey = Object.getOwnPropertyNames( sendData[0] )
// console.log(sendKey);       // [ 0:'__rowNum__', 1:'mail', 2:'fileName' ]


let endData = {};
let inData = [];
let inArr = {}
const sendLen = sendData.length;
let countSend = 0;

var sendMails = new Promise(function(resolve, reject){
    for( var line in sendData ) {
        var sender = '보내는 사람 < cotostar09@gmail.com >';    //보내는 메일
        var receiver = sendData[line][sendKey[1]];             //받는 메일
        var mailTitle = '제목 1';
        var html = '<h1>내용</h1> 본문1';
        var attachments = [
                {
                    //fileName: 'cth_main.jpg',
                    //streamSource: fs.createReadStream('./upload/cth_main.jpg'),
                    path: './upload/'+ sendData[line][sendKey[2]] + '.pdf'
    
                }
            ];
    
        var mailOptions = {
        from: sender,
        to: receiver,
        subject: mailTitle,
        html: html,
        attachments: attachments
        };
    
        var transporter = nodemailer.createTransport(smtpPool({
        service: config.mailer.service,
        host: config.mailer.host,
        port: config.mailer.port,
        auth: {
            user: config.mailer.user,
            pass: config.mailer.password
        },
        tls: {
            rejectUnauthorize: false
        },
        maxConnections: 5,
        maxMessages: 10
        }));
    
        transporter.sendMail(mailOptions, function (err, res) {
        if (err) {
            console.log('failed E-mail : '+mailOptions.to+ ' / ERR => ' + err);
            console.log(res);
            inArr['check'] = 'failed';
            //inArr['email'] = res.accepted ;
            countSend++;
        } else {
            console.log('succeed E-mail : '+ res.accepted + ' / ' + res.response );
            //console.log(res );
            inArr['check'] = 'succeed';
            inArr['email'] = res.accepted ;
            countSend++;
        }
        inData[line] = inArr;
        transporter.close();
        });
    }
    
    return console.log('NEXT');
});

sendMails.then(res => {
    endData['Sheet1'] =  inData;
    let finalHeaders = ['email', 'check'];
    
    //엑셀 저장
    let ws = XLSX.utils.json_to_sheet(endData['Sheet1'], {header: finalHeaders});
    let wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SheetJS")
    let exportFileName = `./upload/out.xlsx`;
    XLSX.writeFile(wb, exportFileName)
    
    console.log('END');
})


