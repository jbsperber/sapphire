// --- Mock data --------------------------------------------------------------
const MOCK_ITEMS = [
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
    source: "Teams",
    title: "Sales standup: pipeline blockers",
    snippet: "Open items for NorthEast region and discount policy.",
    link: "https://teams.microsoft.com/"
  }
];

// --- Adaptive Cards ---------------------------------------------------------
const inputCard = {
  type: "AdaptiveCard",
  $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
  version: "1.5",
  body: [
    { type: "TextBlock", text: "Search Teams & Outlook (mocked)", weight: "Bolder", size: "Medium" },
    { type: "TextBlock", text: "Enter keywords and click Search", isSubtle: true, wrap: true },
    { type: "Input.Text", id: "q", placeholder: "e.g., Q4 forecast", isRequired: true }
  ],
  actions: [
    { type: "Action.Submit", title: "Search", data: { verb: "search" } }
  ]
};

function buildResultsCard(query, items) {
  return {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.5",
    body: [
      { type: "TextBlock", text: `Results for: "${query}"`, weight: "Bolder", size: "Medium", wrap: true },
      ...(items.length === 0
        ? [{ type: "TextBlock", text: "No matches found.", isSubtle: true }]
        : items.flatMap(hit => ([
            { type: "TextBlock", text: `• ${hit.source}: ${hit.title}`, weight: "Bolder", wrap: true },
            { type: "TextBlock", text: hit.snippet, isSubtle: true, wrap: true },
            { type: "ActionSet", actions: [{ type: "Action.OpenUrl", title: "Open", url: hit.link }] },
            { type: "TextBlock", text: "", spacing: "Small" }
          ])))
    ]
  };
}

/**
 * Registers mock search behavior on your Teams AI Library `app` instance.
 * @param {import('@microsoft/teams-ai').Application} app
 */
function registerMockSearch(app) {
  // 1) On any message → show the input Adaptive Card
  app.on('message', async ({ send, activity }) => {
    const attachment = {
      contentType: "application/vnd.microsoft.card.adaptive",
      content: inputCard
    };
    
    await send({ attachments: [attachment] });
  });

  // 2) Handle Adaptive Card Action.Execute ("search")
  app.on('message.submit.action', async ({ send, activity }) => {
    if (activity.value && activity.value.verb === 'search') {
      const q = (activity.value.q ?? "").trim();
      const filtered = MOCK_ITEMS.filter(x =>
        (x.title + " " + x.snippet).toLowerCase().includes(q.toLowerCase())
      );

      // Send a results card as a new message
      const attachment = {
        contentType: "application/vnd.microsoft.card.adaptive",
        content: buildResultsCard(q, filtered)
      };
      
      await send({ attachments: [attachment] });
    }
  });
}

module.exports = { registerMockSearch };