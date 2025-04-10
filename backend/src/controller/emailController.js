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
      <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">Hi Sagar,</p>
      <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;"><strong style="color: #0a7dda;">Warm greetings from Hoping Minds.</strong></p>
      <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">
        We are excited to extend an invitation for <strong>On-Campus & Off-Campus Hiring</strong> through 
        <strong style="color: #0a7dda;">Hoping Minds</strong>, offering you access to a diverse pool of highly skilled, 
        job-ready talent — <span style="color: green;"><strong>at zero cost</strong></span>.
        We would love to connect and understand your hiring needs and explore how our trained graduates can add value to your organization.
      </p>

      <h3 style="font-family: Arial, sans-serif; color: #0a7dda;">Why Partner with Hoping Minds?</h3>
      <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">
        <strong>Hoping Minds</strong> runs <strong>Industry-Oriented Programs</strong> designed to equip students with 
        hands-on experience, corporate readiness, and holistic development. 
        Our talent pool is <strong>rigorously trained</strong> and ready to contribute from <strong>Day 1</strong>.
      </p>
      <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;"><strong>We offer skilled candidates across both Technical and Business Functional domains, including:</strong></p>
      <ul style="font-family: Arial, sans-serif; font-size: 16px; color: #333; padding-left: 20px;">
        <li><strong>Technical:</strong> Data Science, Full Stack Development, Electric Vehicle Design, Hydrocarbon, AWS, Cybersecurity, and more</li>
        <li><strong>Business Functions:</strong> Sales, Marketing, Human Resources, Finance, Business Operations, Customer Support, and more</li>
      </ul>
      <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;"><strong>Our dynamic curriculum aligns with real-time industry demands, covering:</strong></p>
      <ul style="font-family: Arial, sans-serif; font-size: 16px; color: #333; padding-left: 20px;">
        <li>Core Domain Knowledge</li>
        <li>Aptitude & Data Interpretation</li>
        <li>Communication & Interview Preparation</li>
        <li>Personality Development & Workplace Etiquette</li>
      </ul>

      <h3 style="font-family: Arial, sans-serif; color: #0a7dda;">Why Top Recruiters Prefer Hoping Minds?</h3>
      <ul style="font-family: Arial, sans-serif; font-size: 16px; color: #333; padding-left: 20px;">
        <li><strong>Streamlined Process:</strong> Access a pre-vetted, diverse talent pool in one go</li>
        <li><strong>Immediate Availability:</strong> Candidates ready for immediate deployment</li>
        <li><strong>Rigorous Talent Selection:</strong> Three-step screening ensures quality</li>
        <li><strong>Job-Ready Workforce:</strong> Trained for both technical and corporate environments</li>
        <li><strong>Cost-Efficient Hiring:</strong> Save on onboarding & training expenses</li>
        <li><strong style="color: green;">Zero Cost to Company:</strong> No commissions, no hidden charges</li>
      </ul>

      <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">
        As a certified <strong>National Skill Development Corporation (NSDC)</strong> Partner, we maintain 
        high-quality training standards. With a growing pool of <strong>1,000+ skilled graduates</strong>, 
        we are confident in offering you tailor-made hiring solutions.
      </p>

      <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">
        We are also associated with top universities and colleges across the country, providing you access 
        to a diverse and pan-India talent base — all through a single platform.
      </p>

      <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">
        We would be delighted to schedule a conversation and discuss how 
        <strong style="color: #0a7dda;">Hoping Minds</strong> can support your recruitment objectives.
      </p>

      <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">Looking forward to hearing from you!</p>
      <br>
      <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">--</p>
      <p style="font-family: Arial, sans-serif; font-size: 16px; color: #333;">
        <strong>Mudit Vigya</strong><br>
        Senior Manager – Placements & Corporate Relations<br>
        <a href="https://www.hopingminds.com" style="color: #0a7dda;">www.hopingminds.com</a><br><br>
        E: <a href="mailto:mudit@hopingminds.com" style="color: #0a7dda;">mudit@hopingminds.com</a><br>
        M: +91 977 988 6900<br>
        A: E-314, 4th Floor, Sector 75, Mohali
      </p>
    `;

    for (const row of data) {
      const { Email, Name, Subject } = row;

      if (!Email || !Name || !Subject) continue;

      const finalMessage = `<p>Dear ${Name},</p>${staticBody}`;

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