import xlsx from 'xlsx';
import fs from 'fs';
import { sendMail } from '../utils/email.js';
import EmailModel from '../models/email.js';

export async function sendEmailsFromExcel(req, res) {
  const filePath = req?.file?.path;

  if (!filePath) {
    return res.status(400).json({ error: 'No file uploaded.' });
  }

  try {
    const workbook = xlsx.readFile(filePath, { raw: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);

    if (!Array.isArray(data) || data.length === 0) {
      fs.unlinkSync(filePath);
      return res.status(400).json({ error: 'Uploaded Excel file is empty or invalid.' });
    }

    let sentCount = 0;

    const staticBody = `
  <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">Dear</p>
  <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;"><strong>Warm greetings from <span style="font-weight: bold;">Hoping Minds</span>.</strong></p>

  <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">
    We are an <strong>EdTech-driven Talent Solutions Partner</strong>, committed to empowering organizations through our comprehensive 
    <span style="color: red;">Recruitment, Training, and Deployment programs</span>. 
    At <span style="font-weight: bold;">Hoping Minds</span>, we aim to bridge the gap between academia and industry by preparing and delivering job-ready talent 
    that aligns with the evolving business landscape.
  </p>

  <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">
    As part of our commitment to excellence, we offer a wide range of services including 
    <span style="color: red;">Campus Recruitment for fresh graduates</span>, 
    <span style="color: red;">Lateral Hiring</span>, 
    <span style="color: red;">Employee Upskill Training</span>, 
    <span style="color: red;">Domain-Specific Training</span>, 
    <span style="color: red;">Corporate Training</span>, 
    <span style="color: red;">Training Needs Identification (TNI)</span> and 
    <span style="color: red;">Training Needs Analysis (TNA)</span>, 
    as well as designing and implementing 
    <span style="color: red;">Performance Matrices</span> 
    to support long-term employee growth and organizational success.
  </p>

  <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">
    We proudly serve as a recruitment partner for renowned companies such as 
    <span style="color: green;">ACME Group</span>, supporting them in both 
    <span style="color: red;">lateral hiring</span> and 
    <span style="color: red;">structured campus drives</span> at prestigious institutions like 
    <span style="color: green;">IITs, NITs, BITs, IIMs</span>, and other leading universities. 
    Our successful placements span across roles including Graduate Engineer Trainees (GETs), Management Trainees, Legal Advisors, and more.
  </p>

  <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">
    Additionally, we have a 
    <span style="color: green;">pool of over 2000+ experienced professionals</span> from relevant domains, ready to add immediate value 
    to forward-thinking organizations like yours.
  </p>

  <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">
    We would be delighted to explore collaboration opportunities with your esteemed organization. 
    Please let us know a convenient time to connect and discuss how 
    <span style="font-weight: bold;">Hoping Minds</span> can support your hiring and talent development goals.
  </p>

  <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">Looking forward to your response.</p>
`;
    const gmailSignature = `
  <br><br>
  <table cellpadding="0" cellspacing="0" style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
    <tr>
      <td style="padding-right: 20px; border-right: 2px solid #73c79f;">
        <p style="margin: 0; font-weight: bold; font-size: 16px; color: #45b47c;">Mudit Vigya</p>
        <p style="margin: 0;">Senior Manager - Placements & Corporate Relations</p>
        <p style="margin: 5px 0;">
          <a href="https://www.hopingminds.com" style="color: #45b47c; text-decoration: none;">www.hopingminds.com</a>
        </p>
        <p style="margin-top: 8px;">
          <a href="https://facebook.com" target="_blank" style="text-decoration: none; margin-right: 5px;">
            <img src="https://cdn-icons-png.flaticon.com/24/733/733547.png" alt="Facebook" style="vertical-align: middle;">
          </a>
          <a href="https://linkedin.com/in/muditvigya" target="_blank" style="text-decoration: none; margin-right: 5px;">
            <img src="https://cdn-icons-png.flaticon.com/24/145/145807.png" alt="LinkedIn" style="vertical-align: middle;">
          </a>
          <a href="https://instagram.com" target="_blank" style="text-decoration: none;">
            <img src="https://cdn-icons-png.flaticon.com/24/2111/2111463.png" alt="Instagram" style="vertical-align: middle;">
          </a>
        </p>
      </td>
      <td style="padding-left: 20px;">
        <a href="https://www.hopingminds.com" target="_blank">
          <img src="https://sbs.ac.in/wp-content/uploads/2023/09/Asset-5.png" alt="Hoping Minds" height="40" style="margin-bottom: 8px;">
        </a><br>
        <p style="margin: 0;"><strong style="color: #d00;">E:</strong> <a href="mailto:mudit@hopingminds.com" style="color: #000;">mudit@hopingminds.com</a></p>
        <p style="margin: 0;"><strong style="color: #d00;">M:</strong> +91 977 988 6900</p>
        <p style="margin: 0;"><strong style="color: #d00;">A:</strong> E-314, 4th Floor, Sector 75, Mohali</p>
      </td>
    </tr>
  </table>
`;
    for (const row of data) {
      const { Email, Name, Subject } = row;

      if (!Email || !Name || !Subject) continue;

      const finalMessage = `<p>Dear ${Name},</p>${staticBody}${gmailSignature}`;

      try {
        await sendMail({
          from: process.env.EMAIL_USER,
          to: Email,
          subject: Subject,
          html: finalMessage,
        });

        await EmailModel.create({
          to: Email,
          subject: Subject,
          message: finalMessage,
        });

        sentCount++;
      } catch (emailError) {
        console.error(`❌ Failed to send email to ${Email}:`, emailError.message);
        continue;
      }
    }

    fs.unlinkSync(filePath); // clean up uploaded file

    return res.status(200).json({ success: true, message: `${sentCount} emails sent.` });
  } catch (error) {
    console.error('❌ Error processing file:', error.message);
    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    return res.status(500).json({ error: 'Failed to send emails from uploaded file.' });
  }
}