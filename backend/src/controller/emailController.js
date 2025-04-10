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

    const staticBody = `
      <p>Hi,</p>
      <p>Warm greetings from HopingMinds.</p>
      <p>We are excited to extend an invitation for On-Campus & Off-Campus Hiring through Hoping Minds, offering you access to a diverse pool of highly skilled, job-ready talent — at zero cost. We would love to connect and understand your hiring needs, and explore how our trained graduates can add value to your organization.</p>

      <h3>Why Partner with Hoping Minds?</h3>
      <p>Hoping Minds runs Industry-Oriented Programs designed to equip students with hands-on experience, corporate readiness, and holistic development. Our talent pool is rigorously trained and ready to contribute from Day 1.</p>
      
      <p><strong>Technical:</strong> Data Science, Full Stack Development, Electric Vehicle Design, Hydrocarbon, AWS, Cybersecurity, and more</p>
      <p><strong>Business Functions:</strong> Sales, Marketing, Human Resources, Finance, Business Operations, Customer Support, and more</p>

      <p>Our dynamic curriculum aligns with real-time industry demands, covering:</p>
      <ul>
        <li>Core Domain Knowledge</li>
        <li>Aptitude & Data Interpretation</li>
        <li>Communication & Interview Preparation</li>
        <li>Personality Development & Workplace Etiquette</li>
      </ul>

      <h3>Why Top Recruiters Prefer Hoping Minds?</h3>
      <ul>
        <li>Streamlined Process: Access a pre-vetted, diverse talent pool in one go</li>
        <li>Immediate Availability: Candidates ready for immediate deployment</li>
        <li>Rigorous Talent Selection: Three-step screening ensures quality</li>
        <li>Job-Ready Workforce: Trained for both technical and corporate environments</li>
        <li>Cost-Efficient Hiring: Save on onboarding & training expenses</li>
        <li>Zero Cost to Company: No commissions, no hidden charges</li>
      </ul>

      <p>As a certified National Skill Development Corporation (NSDC) Partner, we maintain high-quality training standards. With a growing pool of 1,000+ skilled graduates, we are confident in offering you tailor-made hiring solutions.</p>

      <p>We are also associated with top universities and colleges across the country, providing you access to a diverse and pan-India talent base—all through a single platform.</p>

      <p>We would be delighted to schedule a conversation and discuss how Hoping Minds can support your recruitment objectives.</p>
      
      <p>Looking forward to hearing from you!</p>
      <br>
      <p>--</p>
      <p><strong>Mudit Vigya</strong><br>
      Senior Manager- Placements & Corporate Relations<br>
      <a href="https://www.hopingminds.com">www.hopingminds.com</a><br>
      <br>
      E: mudit@hopingminds.com<br>
      M: +91 977 988 6900<br>
      A: E-314, 4th Floor, Sector 75, Mohali</p>
    `;

    for (const row of data) {
      const { Email, Name, Subject } = row;
      if (!Email || !Name || !Subject) continue;

      const finalMessage = `<p>Dear ${Name},</p>${staticBody}`;

      await sendMail({
        from: process.env.EMAIL_USER,
        to: Email,
        subject: Subject,
        html: finalMessage,
      });

      await email.create({
        to: Email,
        subject: Subject,
        message: finalMessage,
      });

      sentCount++;
    }

    fs.unlinkSync(filePath); // delete uploaded file after processing
    res.status(200).json({ success: true, message: `${sentCount} emails sent.` });
  } catch (error) {
    console.error('❌ Error processing file:', error);
    res.status(500).json({ error: 'Failed to send emails from uploaded file.' });
  }
}