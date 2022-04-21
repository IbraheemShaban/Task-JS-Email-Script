// Your code goes here

'use strict';
const excelToJson = require('convert-excel-to-json');
const result = excelToJson({
  sourceFile: 'names.xlsx',
});

const today = new Date();
const day = String(today.getDate()).padStart(2, '0');
const month = String(today.getMonth() + 1).padStart(2, '0');
const year = today.getFullYear();

const date = `issued ${day}/${month}/${year}`;

result['Sheet1'].forEach((e) => {
  const sgMail = require('@sendgrid/mail');

  let fileContent = `<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    
    <title>Certificate </title>
</head>

<body background="http://cdn.mcauto-images-production.sendgrid.net/eb94cb5adc0068e9/9a1b041d-5a23-48b8-b689-7c4f42f9fc11/800x1200.jpeg" width="50px" height="50px">
    <div style="color:black;text-align:center;font-family: 'Trebuchet MS', sans-serif;">
        <h1><b>CERTIFICATE OF COMPLETION</b></h1>
        <p>This certifies that</p>
        <h2>${e['A']}</h2>
        <p>has completed and passed the course in</p>
        <h2>${e['C']}</h2>
        <p>with a grade of</p>
        <h2>${e['D']}</h2>
        <p>${date}</p>
    </div>
</body>
<script src="/app.js"></script>

</html>`;

  sgMail.setApiKey(process.env.SENDGRID_API_KEY);
  const msg = {
    to: e['B'], // Change to your recipient
    from: 'ish3ban1@gmail.com', // Change to your verified sender
    subject: 'Certificate',
    text: 'Certificate',
    html: fileContent,
  };
  sgMail
    .send(msg)
    .then(() => {
      console.log('Email sent');
    })
    .catch((error) => {
      console.error(error);
    });
});
