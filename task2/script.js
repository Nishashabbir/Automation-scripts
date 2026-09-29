function checkWebsites() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const START_ROW = 2; //when you rerun it , everytime it will start from that row

  const PHONE_COL = 3;        // C
  const WEBSITE_COL = 4;      // D

  const HAS_WEBSITE_COL = 5;  // E
  const HAS_WHATSAPP_COL = 6; // F
  const CHAT_WIDGET_COL = 7;  // G

  const BATCH_SIZE = 30;

  const lastRow = sheet.getLastRow();

  if (lastRow < START_ROW) {
    Logger.log("No records found.");
    return;
  }


  // ==========================================
  // GET LAST PROCESSED ROW
  // ==========================================

  const properties =
    PropertiesService.getScriptProperties();

  let nextRow =
    Number(
      properties.getProperty("NEXT_ROW")
    );

  if (!nextRow || nextRow < START_ROW) {
    nextRow = START_ROW;
  }


  // ==========================================
  // IF EVERYTHING IS FINISHED
  // ==========================================

  if (nextRow > lastRow) {

    Logger.log("All records are already processed.");

    properties.deleteProperty("NEXT_ROW");

    return;
  }


  // ==========================================
  // PROCESS ONE BATCH
  // ==========================================

  const batchSize =
    Math.min(
      BATCH_SIZE,
      lastRow - nextRow + 1
    );


  Logger.log(
    "Starting rows " +
    nextRow +
    " - " +
    (nextRow + batchSize - 1)
  );


  const data =
    sheet
      .getRange(
        nextRow,
        1,
        batchSize,
        4
      )
      .getValues();


  const urls = [];
  const phones = [];


  data.forEach(row => {

    urls.push(
      normalizeWebsite(
        row[WEBSITE_COL - 1]
      )
    );

    phones.push(
      normalizePhone(
        row[PHONE_COL - 1]
      )
    );

  });


  // ==========================================
  // FETCH WEBSITES
  // ==========================================

  const responses =
    fetchWebsitesSafely(urls);


  const output = [];


  // ==========================================
  // ANALYZE
  // ==========================================

  for (let i = 0; i < batchSize; i++) {

    const url = urls[i];

    const response =
      responses[i];


    let hasWebsite = false;
    let hasWhatsApp = false;
    let hasChatWidget = false;


    // No URL
    if (!url) {

      output.push([
        false,
        false,
        false
      ]);

      continue;
    }


    // Request failed
    if (!response) {

      Logger.log(
        "Could not access: " + url
      );

      output.push([
        false,
        false,
        false
      ]);

      continue;
    }


    const statusCode =
      response.getResponseCode();


    let html = "";

    try {

      html =
        response
          .getContentText()
          .toLowerCase();

    } catch (error) {

      html = "";
    }


    // ========================================
    // WEBSITE
    // ========================================

    hasWebsite =
      statusCode >= 200 &&
      statusCode < 400;


    // ========================================
    // WHATSAPP + CHAT
    // ========================================

    if (hasWebsite) {

      hasWhatsApp =
        detectWhatsApp(
          html,
          phones[i]
        );


      hasChatWidget =
        detectChatWidget(html);
    }


    output.push([
      hasWebsite,
      hasWhatsApp,
      hasChatWidget
    ]);

  }


  // ==========================================
  // WRITE RESULTS
  // ==========================================

  sheet
    .getRange(
      nextRow,
      HAS_WEBSITE_COL,
      output.length,
      3
    )
    .setValues(output);


  // ==========================================
  // SAVE CHECKPOINT
  // ==========================================

  const nextStart =
    nextRow + batchSize;


  properties.setProperty(
    "NEXT_ROW",
    String(nextStart)
  );


  Logger.log(
    "Processed rows " +
    nextRow +
    " - " +
    (nextStart - 1)
  );


  if (nextStart <= lastRow) {

    Logger.log(
      "Next run will start at row " +
      nextStart
    );

  } else {

    Logger.log(
      "ALL RECORDS PROCESSED."
    );

    properties.deleteProperty("NEXT_ROW");
  }
}


// =====================================================
// SAFE WEBSITE FETCHING
// =====================================================

function fetchWebsitesSafely(urls) {

  const results =
    new Array(urls.length).fill(null);


  const validIndexes = [];


  urls.forEach((url, index) => {

    if (url) {
      validIndexes.push(index);
    }

  });


  if (validIndexes.length === 0) {
    return results;
  }


  /*
   * Try multiple websites simultaneously.
   */
  const requests = validIndexes.map(index => {

    return {
      url: urls[index],
      method: "get",
      muteHttpExceptions: true,
      followRedirects: true
    };

  });


  try {

    const responses =
      UrlFetchApp.fetchAll(requests);


    responses.forEach((response, i) => {

      results[validIndexes[i]] = response;

    });


    return results;

  } catch (error) {

    /*
     * If one website causes fetchAll to fail,
     * check them individually instead.
     */

    Logger.log(
      "Batch fetch issue: " +
      error.message
    );


    validIndexes.forEach(index => {

      try {

        results[index] =
          UrlFetchApp.fetch(
            urls[index],
            {
              method: "get",
              muteHttpExceptions: true,
              followRedirects: true
            }
          );

      } catch (error) {

        Logger.log(
          "Website failed: " +
          urls[index]
        );

        results[index] = null;
      }

    });

  }


  return results;
}



// =====================================================
// NORMALIZE WEBSITE
// =====================================================

function normalizeWebsite(value) {

  if (!value) return "";

  let website =
    String(value).trim();


  if (!website) {
    return "";
  }


  // Remove spaces
  website =
    website.replace(/\s+/g, "");


  // Add HTTPS if protocol missing
  if (!/^https?:\/\//i.test(website)) {

    website =
      "https://" + website;

  }


  return website;
}



// =====================================================
// NORMALIZE PHONE
// =====================================================

function normalizePhone(value) {

  if (!value) return "";

  return String(value)
    .trim()
    .replace(/[^\d+]/g, "");
}



// =====================================================
// WHATSAPP DETECTION
// =====================================================

function detectWhatsApp(html, phone) {

  // Direct WhatsApp links
  if (
    html.includes("wa.me/") ||
    html.includes("api.whatsapp.com") ||
    html.includes("whatsapp.com/send") ||
    html.includes("whatsapp://")
  ) {

    return true;
  }


  // WhatsApp-related text
  if (
    html.includes("whatsapp us") ||
    html.includes("chat on whatsapp") ||
    html.includes("message us on whatsapp") ||
    html.includes("whatsapp chat")
  ) {

    return true;
  }


  // Check supplied phone number
  // appearing in a WhatsApp URL
  if (phone) {

    const digits =
      phone.replace(/\D/g, "");


    if (digits.length >= 7) {

      if (
        html.includes("wa.me/" + digits) ||
        html.includes("phone=" + digits)
      ) {

        return true;
      }
    }
  }


  return false;
}



// =====================================================
// CHAT WIDGET DETECTION
// =====================================================

function detectChatWidget(html) {

  const patterns = [

    // Intercom
    "intercom",

    // Tawk.to
    "tawk.to",

    // Drift
    "drift.com",
    "drift.js",

    // HubSpot
    "hubspot",
    "hs-scripts",

    // Zendesk
    "zendesk",
    "zopim",

    // Crisp
    "crisp.chat",

    // LiveChat
    "livechatinc",

    // Tidio
    "tidio",

    // Freshchat
    "freshchat",

    // Smartsupp
    "smartsupp",

    // Chatra
    "chatra",

    // Olark
    "olark",

    // Facebook Messenger
    "customerchat",

    // Generic widgets
    "chat-widget",
    "chatwidget",
    "chat_widget",
    "live-chat-widget",
    "livechat-widget",
    "chatbot-widget"
  ];


  return patterns.some(
    pattern => html.includes(pattern)
  );
}