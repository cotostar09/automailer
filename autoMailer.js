const nodemailer = require('nodemailer');
const smtpPool = require('nodemailer-smtp-pool');
const config = require('./config/config.json');
const fs = require('fs');
const XLSX = require("xlsx");

//database conn
const mongojs = require('mongojs');
const db = mongojs('autoMailCheck', ['checkSendOk']);

 
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

for( var line in sendData ){
    //DB 초기화
    db.checkSendOk.remove({});
    //DB 데이터 입력
    db.checkSendOk.save({mail:sendData[line][sendKey[1]], fileName:sendData[line][sendKey[2]]});
}

var sendMails = new Promise(function(resolve, reject){
    for( var line in sendData ) {
        var sender = '서던포스트 < ysj@southernpost.co.kr >';    //보내는 메일
        
       //var sender = '보내는 사람 < cotostar09@gmail.com >';    //보내는 메일 
        var receiver = sendData[line][sendKey[1]];             //받는 메일
        var mailTitle = '영등포구청_간부진 부패위험성 진단 평가 결과 송부';
        var html = '<div>안녕하십니까</div><div><br></div><div>영등포구청 감사실 의뢰로 지난 6월 &lt;간부진 부패위험성 진단 평가&gt;를 실시한 평가결과를 붙임과 같이 첨부하여 보내드립니다.</div><div><br></div><div>본 평가는 국민권익위원회의 고위공직자 부패위험성 진단 표준모형을 준수하였으며, 영등포구청 간부진 국/소장, 부서장/동장 총 62명을 대상으로 조사기관 (주)서던포스트에서 진행하였습니다.</div><div><br></div><div>파일첨부:</div><div>간부진 청렴도 인사말 및 개인통보용 결과 안내</div>';
        var attachments = [
                {
                    //fileName: 'cth_main.jpg',
                    //streamSource: fs.createReadStream('./upload/cth_main.jpg'),
                    path: './upload/'+ sendData[line][sendKey[2]] + '.pdf'
                },
                {
                    //fileName: 'cth_main.jpg',
                    //streamSource: fs.createReadStream('./upload/cth_main.jpg'),
                    path: './upload/(첨부) 간부진 청렴도 인사말.hwp'
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
            countSend++;
            console.log(countSend + '. failed E-mail : '+mailOptions.to+ ' / ERR => ' + err);
            console.log(err.path.split('\\').pop().split('.')[0]);
            var errFileName = err.path.split('\\').pop().split('.')[0];
            
            db.checkSendOk.update({fileName:errFileName}, {$set:{check: 'failed'}});

        } else {
            countSend++;
            console.log(countSend + '. succeed E-mail : '+ res.accepted + ' / ' + res.response );
            console.log(String(res.accepted));
            inArr['check'] = 'succeed';
            inArr['email'] = res.accepted ;
            db.checkSendOk.update({mail:String(res.accepted)}, {$set:{check: 'succeed'}});
        }
        inData[line] = inArr;
        transporter.close();
        });
    }
    
    return console.log('NEXT');
});

sendMails.then(res => {

})


