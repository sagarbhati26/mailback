// bulkSend.js
import { sendMail } from './sendMail.js';
import pLimit from 'p-limit';
import fs from 'fs';
import { createObjectCsvWriter } from 'csv-writer';

// Set concurrency limit
const limit = pLimit(5); // 5 emails at a time (adjustable)

const bulkSendEmails = async (emailList) => {
  const failedEmails = [];

  // Wrap each email sending inside limit()
  const tasks = emailList.map(({ to, subject, html }) =>
    limit(async () => {
      const result = await sendMail({ to, subject, html });
      if (!result.success) {
        failedEmails.push({ email: to, reason: result.error });
      }
    })
  );

  // Wait for all tasks to complete
  await Promise.all(tasks);

  // Save failed ones to CSV
  if (failedEmails.length > 0) {
    const csvWriter = createObjectCsvWriter({
      path: 'failed_emails.csv',
      header: [
        { id: 'email', title: 'Email' },
        { id: 'reason', title: 'Reason' },
      ],
    });

    await csvWriter.writeRecords(failedEmails);
    console.log(`📄 Saved ${failedEmails.length} failed emails to failed_emails.csv`);
  } else {
    console.log('✅ All emails sent successfully!');
  }
};

export { bulkSendEmails };