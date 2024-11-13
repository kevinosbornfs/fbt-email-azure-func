import * as msal from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { config } from "../config";

const cca = new msal.ConfidentialClientApplication({
  auth: {
    clientId: config.clientId,
    authority: `${config.authority}/${config.tenantId}`,
    clientSecret: config.clientSecret
  }
});

async function getAccessToken() {
  const result = await cca.acquireTokenByClientCredential({
    scopes: config.scope,
  });
  return result.accessToken;
}

function getGraphClient(accessToken: string) {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    }
  });
}

const sendDailyEmail = async function (context: any): Promise<void> {  
  const today = new Date();
  const dayOfWeek = today.getUTCDay(); 

  if (dayOfWeek >= 1 && dayOfWeek <= 5) {
    const to = "recipient@example.com";
    const subject = "Daily Reminder";
    const body = "This is your daily reminder email sent using Microsoft Graph API.";

    try {
      const accessToken = await getAccessToken();
      const graphClient = getGraphClient(accessToken);

      const message = {
        subject: subject,
        body: {
          contentType: "Text",
          content: body
        },
        toRecipients: [
          {
            emailAddress: {
              address: to
            }
          }
        ]
      };

      await graphClient.api("/me/sendMail").post({ message });

      context.log(`Email sent successfully to ${to}`);
    } catch (error) {
      context.log(`Error sending email: ${error.message}`);
    }
  } else {
    context.log("Today is a weekend. No email sent.");
  }
};

export default sendDailyEmail;
