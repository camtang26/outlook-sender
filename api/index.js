// Add dotenv config
require('dotenv').config();

const express = require("express");
const axios = require("axios");
const { ClientSecretCredential } = require("@azure/identity");
const serverless = require("serverless-http");

const app = express();
app.use(express.json());

// 1. Environment variables (in Vercel or local .env):
//    AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, SENDER_UPN
const tenantId = process.env.AZURE_TENANT_ID;
const clientId = process.env.AZURE_CLIENT_ID;
const clientSecret = process.env.AZURE_CLIENT_SECRET;

// The mailbox from which you want to send (the user principal name or user ID)
const senderUpn = process.env.SENDER_UPN; // e.g., "no-reply@mytenant.onmicrosoft.com"

// 2. Create a credential object from the Azure Identity library
const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

// 3. POST endpoint to send email
app.post("/send-email", async (req, res) => {
  try {
    const { recipient, subject, body } = req.body;

    if (!recipient || !subject || !body) {
      return res.status(400).json({ error: "Missing required fields." });
    }

    // Get an access token for MS Graph
    const tokenResponse = await credential.getToken("https://graph.microsoft.com/.default");
    const accessToken = tokenResponse.token;

    // Graph sendMail endpoint for application permissions:
    // POST /users/{senderUpn}/sendMail
    const endpoint = `https://graph.microsoft.com/v1.0/users/${senderUpn}/sendMail`;

    // Prepare the email payload
    const emailData = {
      message: {
        subject: subject,
        body: {
          contentType: "HTML",
          content: body
        },
        toRecipients: [
          {
            emailAddress: {
              address: recipient
            }
          }
        ]
      },
      saveToSentItems: "true"
    };

    // Make the POST request to Graph
    await axios.post(endpoint, emailData, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      timeout: 20000
    });

    return res.json({ message: "Email sent successfully!" });
  } catch (error) {
    console.error(error?.response?.data || error);
    return res.status(500).json({ error: "Failed to send email." });
  }
});

// Server start for local development or export for serverless (Vercel) deployment
if (require.main === module) {
  const port = process.env.PORT || 3000;
  app.listen(port, () => {
    console.log(`Server listening on port ${port}`);
  });
} else {
  module.exports = serverless(app, { binary: false });
} 