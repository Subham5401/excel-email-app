const excelJs = require('exceljs');
const nodemailer = require('nodemailer');

const transporter = nodemailer.createTransport({
  service: 'outlook',
  auth: {
    user: process.env.EMAIL,
    pass: process.env.EMAIL_PASSWORD,
  },
});

const emailSign = `
    Best regards,
    Subham Mishra
`;

const sendEmailsFromExcel = async () => {
  const workbook = new excelJs.Workbook();

  await workbook.xlsx.readFile('');

  const worksheet = workbook.getWorksheet(1);

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      const email = row.getCell(1).value;
      const subject = row.getCell(2).value;
      const body = row.getCell(3).value;

      const bodyWithSign = `${body}\n\n${emailSign}`;

      const mailOptions = {
        from: process.env.EMAIL,
        to: email,
        subject: subject,
        text: bodyWithSign,
      };

      transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
          console.log(`Error sending email to ${email}: `, error);
        } else {
          console.log(`Email sent to ${email}: ${info.response}`);
        }
      });
    }
  });
};

sendEmailsFromExcel();
