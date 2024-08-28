const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');
const bodyParser = require('body-parser');
require('dotenv').config();
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');

const app = express();
const port = process.env.PORT || 3000;

// CORS configuration using the cors package
app.use(cors({
  origin: 'http://localhost:3000', // Allow requests from this origin
  methods: ['GET', 'POST'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));

app.use(bodyParser.json());

// Google Sheets and Drive credentials
const oauth2Client = new google.auth.OAuth2(
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET,
  process.env.REDIRECT_URI
);

oauth2Client.setCredentials({ refresh_token: process.env.REFRESH_TOKEN });

const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
const drive = google.drive({ version: 'v3', auth: oauth2Client });
const slides = google.slides({ version: 'v1', auth: oauth2Client });

// Nodemailer configuration
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: 'shivanshraheja81@gmail.com',
    pass: 'fvzi qlun sppd hszv' // Be sure to keep this secure
  }
});

// Log file path
const logFilePath = path.join(__dirname, 'logs.json');

// Route to generate certificates
app.post('/generate-certificates', async (req, res) => {
  const { sheetId, sheetName, webinarName, date, organizedBy } = req.body;

  if (!sheetId || !sheetName || !webinarName || !date || !organizedBy) {
    return res.status(400).json({ status: 'error', message: 'One or more parameters are missing.' });
  }

  try {
    const sheetData = await getSheetData(sheetId, sheetName);
    const totalCertificates = sheetData.length - 1;
    let certificatesGenerated = 0;

    await generateCertificates(sheetData, webinarName, date, organizedBy, (generatedCount) => {
      certificatesGenerated = generatedCount;
    });

    // Update log
    const logData = { totalCertificates, certificatesGenerated };
    fs.writeFileSync(logFilePath, JSON.stringify(logData));

    res.json({ status: 'success', message: 'Certificates generated successfully!', totalCertificates, certificatesGenerated });
  } catch (error) {
    console.error('Error in /generate-certificates:', error);
    res.status(500).json({ status: 'error', message: 'An error occurred while generating certificates. Please check the server logs.' });
  }
});

// Route to fetch logs
app.get('/fetch-logs', (req, res) => {
  if (fs.existsSync(logFilePath)) {
    const logData = JSON.parse(fs.readFileSync(logFilePath, 'utf8'));
    res.json(logData);
  } else {
    res.status(404).json({ status: 'error', message: 'Log file not found.' });
  }
});

// Function to get data from Google Sheets
async function getSheetData(sheetId, sheetName) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: sheetName,
  });
  return response.data.values;
}

// Function to generate certificates
async function generateCertificates(sheetData, webinarName, date, organizedBy, updateGeneratedCount) {
  if (!Array.isArray(sheetData) || sheetData.length === 0) {
    throw new Error('No data found in the Google Sheet.');
  }

  const templateId = process.env.TEMPLATE_ID;
  const folderId = process.env.FOLDER_ID;
  let generatedCount = 0;

  for (let i = 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    const name = row[0]?.toString() || '';
    const schoolName = row[2]?.toString().toUpperCase() || '';
    const email = row[1]?.toString() || '';
    const certificateNumber = row[3]?.toString().toUpperCase() || '';

    if (!name || !schoolName || !email || !certificateNumber) {
      console.log(`Skipping row ${i + 1} due to missing data.`);
      continue;
    }

    console.log(`Generating certificate ${i} of ${sheetData.length - 1} for ${name}...`);
    const formattedDate = formatDateToReadable(new Date(date));

    const copyFile = await drive.files.copy({
      fileId: templateId,
      requestBody: {
        name: `${name} - Certificate`,
        parents: [folderId]
      }
    });

    const copyId = copyFile.data.id;

    await slides.presentations.batchUpdate({
      presentationId: copyId,
      requestBody: {
        requests: [
          { replaceAllText: { containsText: { text: '{{Name}}' }, replaceText: name } },
          { replaceAllText: { containsText: { text: '{{SchoolName}}' }, replaceText: schoolName } },
          { replaceAllText: { containsText: { text: '{{WebinarName}}' }, replaceText: webinarName } },
          { replaceAllText: { containsText: { text: '{{Date}}' }, replaceText: formattedDate } },
          { replaceAllText: { containsText: { text: '{{OrganizedBy}}' }, replaceText: organizedBy } },
          { replaceAllText: { containsText: { text: '{{CERT-NUMBER}}' }, replaceText: certificateNumber } }
        ],
      },
    });

    const exportUrl = `https://www.googleapis.com/drive/v3/files/${copyId}/export?mimeType=application/pdf`;
    const response = await drive.files.export({
      fileId: copyId,
      mimeType: 'application/pdf',
    }, { responseType: 'stream' });

    const filename = `${name}_${certificateNumber}.pdf`;

    await sendEmailWithAttachment(
      email,
      `Luneblaze certificate for the session on ${webinarName}`,
      `Dear Educator,<br><br>
       Greetings of the day!!<br><br>
       Hope you are doing well.<br><br>
       This email is to acknowledge your participation in the <b>${webinarName}</b> Session held on <b>${date}</b>, organised by Luneblaze. Please find your Participation Certificate attached.<br><br>
       We organise sessions focusing on SQAAF every month.<br><br>
       Luneblaze is also helping 100+ schools in their SQAAF Journey by assisting in documentation, implementation and self-assessment.<br><br>
       We would like to discuss the possibility of helping your esteemed institution in the SQAAF Implementation journey.<br><br>
       For more details reach out to us at: <b>+91 7533051785</b><br><br>
       Looking forward to the opportunity to support your accreditation needs.<br><br>
       PFA
       Best Regards,
       Team Luneblaze`,
      response.data,
      filename
    );
    
    await drive.files.update({
      fileId: copyId,
      requestBody: { trashed: true }
    });

    generatedCount++;
    console.log(`Certificate ${generatedCount} generated out of ${sheetData.length - 1} for ${name} and sent via email.`);
  }

  updateGeneratedCount(generatedCount); // Callback to update generated count
}

// Function to format date to a readable format
function formatDateToReadable(date) {
  const monthNames = [
    "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
    "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"
  ];

  const day = date.getDate();
  const year = date.getFullYear();
  const month = monthNames[date.getMonth()];

  let suffix = "th";
  if (day === 1 || day === 21 || day === 31) {
    suffix = "st";
  } else if (day === 2 || day === 22) {
    suffix = "nd";
  } else if (day === 3 || day === 23) {
    suffix = "rd";
  }

  return `${month} ${day}${suffix}, ${year}`;
}

// Function to send email with PDF attachment directly from the stream
async function sendEmailWithAttachment(to, subject, htmlContent, pdfStream, filename) {
  const mailOptions = {
    from: 'shivanshraheja81@gmail.com',
    to: to,
    subject: subject,
    html: htmlContent,
    attachments: [
      {
        filename: filename,
        content: pdfStream,
        encoding: 'base64',
        contentType: 'application/pdf'
      }
    ]
  };

  await transporter.sendMail(mailOptions);
}

// Start server
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
