// sendMail.js
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
      html,
    };

    const info = await transporter.sendMail(mailOptions);
    console.log(`📧 Email sent to ${to}:`, info.response);

    // Return success so the bulk processor knows
    return { to, subject, success: true };
  } catch (error) {
    console.error(`❌ Error sending email to ${to}:`, error.message);

    // Return failure info
    return { to, subject, success: false, error: error.message };
  }
};

export { sendMail };