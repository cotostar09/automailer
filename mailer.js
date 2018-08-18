var nodemailer = require('nodemailer');
var smtpPool = require('nodemailer-smtp-pool');
var config = require('./config/config.json');
var fs = require('fs');

var sender = '보내는 사람 < cotostar09@gmail.com >';    //보내는 메일
var receiver = 'cotostar09@gmail.com';                 //받는 메일
var mailTitle = '메일 첨부 파일 테스트3';
var html = '<h1>HTML TEST</h1>';
var attachments = [
        {
            //fileName: 'cth_main.jpg',
            //streamSource: fs.createReadStream('./upload/cth_main.jpg'),
			path: './upload/cth_main.jpg'

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
    console.log('failed... => ' + err);
  } else {
    console.log('succeed... => ' + res.message);
  }
  transporter.close();
});

