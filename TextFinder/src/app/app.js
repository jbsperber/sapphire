const { ManagedIdentityCredential } = require("@azure/identity");
const { App } = require("@microsoft/teams.apps");
const { LocalStorage } = require("@microsoft/teams.common");
const config = require("../config");

// Create storage for conversation history
const storage = new LocalStorage();

const createTokenFactory = () => {
  return async (scope, tenantId) => {
    const managedIdentityCredential = new ManagedIdentityCredential({
        clientId: process.env.CLIENT_ID
      });
    const scopes = Array.isArray(scope) ? scope : [scope];
    const tokenResponse = await managedIdentityCredential.getToken(scopes, {
      tenantId: tenantId
    });
   
    return tokenResponse.token;
  };
};

// Configure authentication using TokenCredentials
const tokenCredentials = {
  clientId: process.env.CLIENT_ID || '',
  token: createTokenFactory()
};

const credentialOptions = config.MicrosoftAppType === "UserAssignedMsi" ? { ...tokenCredentials } : undefined;

// Create the app with storage
const app = new App({
  ...credentialOptions,
  storage
});

// Handle incoming messages with mock search
app.on('message', async ({ send, activity }) => {
  try {
    console.log('Received message:', activity.text);
    const query = activity.text.trim();
    
    // Mock data
    const mockItems = [
      {
        source: "Teams",
        title: "Q4 forecast sync notes",
        snippet: "We agreed to revise the top-line by 8% and regroup Friday.",
        link: "https://teams.microsoft.com/"
      },
      {
        source: "Outlook",
        title: "Re: Q4 forecast spreadsheet", 
        snippet: "Attached the latest workbook with comments.",
        link: "https://outlook.office.com/"
      },
      {
        source: "Outlook",
        title: "Re: Silent install of Mastercam 2027", 
        snippet: "Run the following command line for silent install.",
        link: "https://outlook.office.com/"
      },
      {
        source: "Teams",
        title: "Sales standup: pipeline blockers",
        snippet: "Open items for NorthEast region and discount policy.",
        link: "https://teams.microsoft.com/"
      },
      {
        source: "Teams",
        title: "Mastercam 2026 Daily is Available",
        snippet: "Product Version: 28.0.7963.0",
        link: "https://teams.microsoft.com/"
      }
    ];

// Filter results based on the message
    const filtered = mockItems.filter(x =>
      (x.title + " " + x.snippet).toLowerCase().includes(query.toLowerCase())
    );

    // Create results message
    let resultsText = `🔍 **Search Results for: "${query}"**\n\n`;
    
    if (filtered.length === 0) {
      resultsText += "❌ No matches found.\n\n";
      resultsText += "Try searching for: **Q4**, **sales**, or **forecast**";
    } else {
      filtered.forEach((item, index) => {
        resultsText += `**${index + 1}. ${item.source}: ${item.title}**\n`;
        resultsText += `${item.snippet}\n`;
        resultsText += `🔗 [Open in ${item.source}](${item.link})\n\n`;
      });
    }

    await send(resultsText);
  } catch (error) {
    console.error('Error in message handler:', error);
    await send("Sorry, there was an error processing your search: " + error.message);
  }
});



app.on('message.submit.feedback', async ({ activity }) => {
  //add custom feedback process logic here
  console.log("Your feedback is " + JSON.stringify(activity.value));
});

module.exports = app;