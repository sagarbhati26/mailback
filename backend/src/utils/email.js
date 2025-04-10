import nodemailer from 'nodemailer';

const sendMail = async ({ to, subject, html }) => {
  try {
    const transporter = nodemailer.createTransport({
      service: 'Gmail',
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS,
      },
      logger: true,
      debug: true,
    });

    const mailOptions = {
      from: process.env.EMAIL_USER,
      to,
      subject,
      html, // 👈 we're sending HTML messages now
    };

    const info = await transporter.sendMail(mailOptions);
    console.log('📧 Email sent: ', info.response);
  } catch (error) {
    console.error('❌ Error sending email:', error);
    throw new Error(`Email could not be sent: ${error.message}`);
  }
};

export { sendMail };