import xlsx from 'xlsx';
import fs from 'fs';
import { sendMail } from '../utils/email.js';
import email from '../models/email.js';

export async function sendEmailsFromExcel(req, res) {
  const filePath = req.file.path;

  try {
    const workbook = xlsx.readFile(filePath, { raw: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);

    let sentCount = 0;

    for (const row of data) {
      const { Email, Name, Message, Subject } = row;

      if (!Email || !Name || !Message || !Subject) continue;

      const personalizedMessage = `
  <div style="font-family: Arial, sans-serif; color: #333; padding: 20px; line-height: 1.6; max-width: 700px; margin: auto;">
    <h2 style="color: #000;">Industry Partnership Proposal – Trained Graduates (Technical & Business Functions Role) @HopingMinds</h2>

    <p>Hi ${Name},</p>

    <p>Warm greetings from <strong>HopingMinds</strong>.</p>

    ${Message} <!-- 🔥 Dynamic content from Excel -->

    <br/>

    <p><strong>Why Partner with Hoping Minds?</strong></p>
    <ul>
      <li>Industry-Oriented Programs with hands-on training</li>
      <li>Pre-vetted, job-ready candidates across tech & business functions</li>
      <li>Zero cost hiring with immediate deployment</li>
    </ul>

    <p>We’d love to schedule a quick call to understand your hiring needs and how we can support you.</p>

    <br/>

    <p><strong>Warm regards,</strong></p>

    <p>
      <strong>Mudit Vigya</strong><br/>
      Senior Manager – Placements & Corporate Relations<br/>
      <a href="https://www.hopingminds.com" target="_blank">www.hopingminds.com</a><br/>
      📧 <a href="mailto:mudit@hopingminds.com">mudit@hopingminds.com</a><br/>
      📞 +91 977 988 6900<br/>
      📍 E-314, 4th Floor, Sector 75, Mohali
    </p>
  </div>
`;

      await sendMail({
        from: process.env.EMAIL_USER,
        to: Email,
        subject: Subject,
        html: personalizedMessage,
      });

      await email.create({
        to: Email,
        subject: Subject,
        message: personalizedMessage,
      });

      sentCount++;
    }

    fs.unlinkSync(filePath); // delete uploaded file after processing

    res.status(200).json({
      success: true,
      message: `${sentCount} emails sent.`,
    });
  } catch (error) {
    console.error('❌ Error processing file:', error);
    res.status(500).json({ error: 'Failed to send emails from uploaded file.' });
  }
}