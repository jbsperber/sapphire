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

// Send a greeting when the chat/conversation is opened
app.on('conversationUpdate', async ({ send, activity }) => {
  try {
    const membersAdded = activity.membersAdded || [];
    const botId = activity.recipient && activity.recipient.id;
    const userAdded = membersAdded.some(member => member.id && member.id !== botId);

    if (userAdded) {
      const cardActivity = {
        type: 'message',
        attachments: [
          {
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: {
              $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
              type: "AdaptiveCard",
              version: "1.4",
              body: [
                {
                  type: "TextBlock",
                  text: "Hello — I'm Sapphire TextFinder",
                  weight: "Bolder",
                  size: "Large"
                },
                {
                  type: "TextBlock",
                  text: "I can search Teams and Outlook for messages and emails. Try one of the quick searches or type your own query.",
                  wrap: true,
                  spacing: "None"
                },
                {
                  type: "TextBlock",
                  text: "Quick suggestions:",
                  weight: "Bolder",
                  separator: true
                },
                {
                  type: "TextBlock",
                  text: "• Q4\n• sales\n• forecast",
                  wrap: true,
                  spacing: "None"
                }
              ],
              actions: [
                {
                  type: "Action.Submit",
                  title: "Search: Q4",
                  data: { __type: "sapphire.search", query: "Q4" }
                },
                {
                  type: "Action.Submit",
                  title: "Search: sales",
                  data: { __type: "sapphire.search", query: "sales" }
                },
                {
                  type: "Action.Submit",
                  title: "Search: forecast",
                  data: { __type: "sapphire.search", query: "forecast" }
                },
                {
                  type: "Action.OpenUrl",
                  title: "Learn more",
                  url: "https://learn.microsoft.com/microsoft-365"
                }
              ]
            }
          }
        ]
      };

      await send(cardActivity);
    }
  } catch (error) {
    console.error('Error in conversationUpdate handler:', error);
  }
});
 
// // Handle incoming messages with mock search
// app.on('message', async ({ send, activity }) => {
//   try {
//     console.log('Received message:', activity.text);
//     const query = activity.text.trim();
    
//     // Mock data
//     const mockItems = [
//       {
//         source: "Teams",
//         title: "Q4 forecast sync notes",
//         snippet: "We agreed to revise the top-line by 8% and regroup Friday.",
//         link: "https://teams.microsoft.com/"
//       },
//       {
//         source: "Outlook",
//         title: "Re: Q4 forecast spreadsheet", 
//         snippet: "Attached the latest workbook with comments.",
//         link: "https://outlook.office.com/"
//       },
//       {
//         source: "Outlook",
//         title: "Re: Silent install of Mastercam 2027", 
//         snippet: "Run the following command line for silent install.",
//         link: "https://outlook.office.com/"
//       },
//       {
//         source: "Teams",
//         title: "Sales standup: pipeline blockers",
//         snippet: "Open items for NorthEast region and discount policy.",
//         link: "https://teams.microsoft.com/"
//       },
//       {
//         source: "Teams",
//         title: "Mastercam 2026 Daily is Available",
//         snippet: "Product Version: 28.0.7963.0",
//         link: "https://teams.microsoft.com/"
//       }
//     ];

// // Filter results based on the message
//     const filtered = mockItems.filter(x =>
//       (x.title + " " + x.snippet).toLowerCase().includes(query.toLowerCase())
//     );

//     // Create results message
//     let resultsText = `🔍 **Search Results for: "${query}"**\n\n`;
    
//     if (filtered.length === 0) {
//       resultsText += "❌ No matches found.\n\n";
//       resultsText += "Try searching for: **Q4**, **sales**, or **forecast**";
//     } else {
//       filtered.forEach((item, index) => {
//         resultsText += `**${index + 1}. ${item.source}: ${item.title}**\n`;
//         resultsText += `${item.snippet}\n`;
//         resultsText += `🔗 [Open in ${item.source}](${item.link})\n\n`;
//       });
//     }

//     await send(resultsText);
//   } catch (error) {
//     console.error('Error in message handler:', error);
//     await send("Sorry, there was an error processing your search: " + error.message);
//   }
// });

// Handle incoming messages with mock search
app.on('message', async ({ send, activity }) => {
  try {
    console.log('Received message:', activity.text);
    const query = (activity.text || '').trim();

    // Mock data
    const mockItems = [
      {
        source: "Teams",
        title: "Q4 forecast sync notes",
        snippet: "We agreed to revise the top-line by 8% and regroup Friday.",
        link: "https://teams.microsoft.com/"
      },
      {
        source: "Teams",
        title: "Quarter Four forecast sync notes",
        snippet: "We agreed to revise the top-line by 8% and regroup Friday for Q4 results.",
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
      },
      {
        source: "Outlook",
        title: "Client: Cambrios proposal",
        snippet: "Attached proposal and contract terms for review.",
        link: "https://outlook.office.com/"
      },
      {
        source: "Teams",
        title: "Onboarding checklist for new hires",
        snippet: "Please complete forms and schedule IT setup.",
        link: "https://teams.microsoft.com/"
      },
      {
        source: "Outlook",
        title: "Invoice INV-1023 - Payment Due",
        snippet: "Reminder: invoice due in 10 days. See attached PDF.",
        link: "https://outlook.office.com/"
      },
      {
        source: "Teams",
        title: "Release notes: Sapphire TextFinder 1.2",
        snippet: "Bug fixes and performance improvements in search ranking.",
        link: "https://teams.microsoft.com/"
      },
      {
        source: "Outlook",
        title: "HR: Benefits enrollment open",
        snippet: "Open enrollment runs from Nov 1-15. Choose your plans.",
        link: "https://outlook.office.com/"
      },
      {
        source: "Teams",
        title: "Marketing campaign: Autumn launch",
        snippet: "Creative briefs and timelines attached for review.",
        link: "https://teams.microsoft.com/"
      },
      {
        source: "Outlook",
        title: "Security patch advisory",
        snippet: "Apply critical updates to company laptops this weekend.",
        link: "https://outlook.office.com/"
      },
      {
        source: "Teams",
        title: "All-hands recording - Sept 2025",
        snippet: "Recording and transcript now available in the channel.",
        link: "https://teams.microsoft.com/"
      },
      {
        source: "Outlook",
        title: "Legal: NDA for vendor",
        snippet: "Please sign and return the attached NDA before onboarding.",
        link: "https://outlook.office.com/"
      },
      {
        source: "Teams",
        title: "Customer feedback: feature requests",
        snippet: "Collected requests from Beta customers and prioritization.",
        link: "https://teams.microsoft.com/"
      }
    ];

    // Try to ask an AI service to find matches. Configure endpoint/key via env:
    // AI_ENDPOINT - POST endpoint that accepts { query, items } and returns JSON.
    // Optionally AI_API_KEY for Authorization: Bearer <key>
    let filtered = [];

    const fetchFn = global.fetch || (() => {
      try { return require('node-fetch'); } catch { return null; }
    })();

    const aiEndpoint = process.env.AI_ENDPOINT;
    if (aiEndpoint && fetchFn) {
      try {
        const payload = { query, items: mockItems };
        const headers = { 'Content-Type': 'application/json' };
        if (process.env.AI_API_KEY) headers['Authorization'] = `Bearer ${process.env.AI_API_KEY}`;

        const res = await fetchFn(aiEndpoint, {
          method: 'POST',
          headers,
          body: JSON.stringify(payload)
        });

        if (res.ok) {
          const body = await res.json();
          // Accept multiple response shapes:
          // - { matches: [0,2] }  (indexes)
          // - { results: [ { ...item }, ... ] } (items)
          if (Array.isArray(body.matches)) {
            filtered = body.matches
              .map(i => mockItems[i])
              .filter(Boolean);
          } else if (Array.isArray(body.results)) {
            filtered = body.results;
          } else {
            // If AI returned a text with JSON inside, try to parse
            filtered = [];
          }
        } else {
          console.warn('AI endpoint returned non-OK status, falling back to local filter:', res.status);
        }
      } catch (aiError) {
        console.warn('AI search failed, falling back to local filter:', aiError.message);
      }
    }

    // Fallback to local filter if AI not used or returned nothing
    if (!filtered || filtered.length === 0) {
      filtered = mockItems.filter(x =>
        (x.title + " " + x.snippet).toLowerCase().includes(query.toLowerCase())
      );
    }

    // Create adaptive card for results
    const cardActivity = {
      type: 'message',
      attachments: [
        {
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: {
            $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
            type: "AdaptiveCard",
            version: "1.4",
            body: [
              {
                type: "TextBlock",
                text: `🔍 Search Results for: "${query}"`,
                weight: "Bolder",
                size: "Large"
              }
            ]
          }
        }
      ]
    };

    if (filtered.length === 0) {
      cardActivity.attachments[0].content.body.push(
        {
          type: "TextBlock",
          text: "❌ No matches found.",
          wrap: true
        },
        {
          type: "TextBlock",
          text: "Try searching for: **Q4**, **sales**, or **forecast**",
          wrap: true
        }
      );
    } else {
      filtered.forEach((item, index) => {
        cardActivity.attachments[0].content.body.push(
          {
            type: "Container",
            items: [
              {
                type: "TextBlock",
                text: `${index + 1}. ${item.source}: ${item.title}`,
                weight: "Bolder",
                wrap: true
              },
              {
                type: "TextBlock",
                text: item.snippet,
                wrap: true
              }
            ],
            separator: true
          }
        );
      });

      // Add action buttons for each result
      cardActivity.attachments[0].content.actions = filtered.map((item, index) => ({
        type: "Action.OpenUrl",
        title: `Open in ${item.source}`,
        url: item.link
      }));
    }

    await send(cardActivity);
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