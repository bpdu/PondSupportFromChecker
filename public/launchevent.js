Office.onReady(() => {});
const DEBUG_FORCE_PROMPT = true;

const DOMAIN_RULES = [
  {
    originalDomain: "@pondiot.com",
    validFromDomain: "@pondiot.com",
    message: "You're replying to a message addressed to a @pondiot.com mailbox. The sender must be from the @pondiot.com domain."
  },
  {
    originalDomain: "@pondmobile.com",
    validFromDomain: "@pondmobile.com",
    message: "You're replying to a message addressed to a @pondmobile.com mailbox. The sender must be from the @pondmobile.com domain."
  }
];

const STRICT_ADDRESS_RULE = {
  originalRecipient: "upsupport@pondiot.com",
  exactFrom: "upsupport@pondiot.com",
  message: "You're replying to a message addressed to upsupport@pondiot.com. The sender must be exactly upsupport@pondiot.com."
};

function onMessageSendHandler(event) {
  if (DEBUG_FORCE_PROMPT) {
    event.completed({
      allowEvent: false,
      errorMessage: "Support From Checker diagnostic: OnMessageSend handler is running."
    });
    return;
  }

  validateReplyFrom().then((result) => {
    if (!result.ok) {
      event.completed({
        allowEvent: false,
        errorMessage: result.message
      });
      return;
    }

    event.completed({ allowEvent: true });
  }).catch((error) => {
    event.completed({
      allowEvent: false,
      errorMessage: "Support From Checker diagnostic: " + String((error && error.message) || error || "unknown error")
    });
  });
}

async function validateReplyFrom() {
  const item = Office.context.mailbox.item;
  const from = await getFromAddress(item);
  if (!from) {
    return { ok: true };
  }

  const inReplyTo = String(item.inReplyTo || "").trim();
  if (!inReplyTo) {
    // New message or unavailable metadata: skip strict reply-chain validation.
    return { ok: true };
  }

  const originalToRecipients = await getOriginalToRecipientsFromEws(inReplyTo);
  if (!originalToRecipients.length) {
    return { ok: true };
  }

  // Most specific rule first.
  if (
    originalToRecipients.includes(STRICT_ADDRESS_RULE.originalRecipient) &&
    from !== STRICT_ADDRESS_RULE.exactFrom
  ) {
    return { ok: false, message: STRICT_ADDRESS_RULE.message };
  }

  for (const recipient of originalToRecipients) {
    for (const domainRule of DOMAIN_RULES) {
      if (!recipient.endsWith(domainRule.originalDomain)) {
        continue;
      }

      if (!from.endsWith(domainRule.validFromDomain)) {
        return { ok: false, message: domainRule.message };
      }
    }
  }

  return { ok: true };
}

function getFromAddress(item) {
  return new Promise((resolve) => {
    item.from.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
        resolve("");
        return;
      }

      resolve(String(result.value.emailAddress || "").toLowerCase());
    });
  });
}

function getOriginalToRecipientsFromEws(inReplyTo) {
  return findOriginalItemIdByInternetMessageId(inReplyTo)
    .then((itemId) => {
      if (!itemId) {
        return [];
      }
      return getToRecipientsByItemId(itemId);
    });
}

function findOriginalItemIdByInternetMessageId(internetMessageId) {
  const escapedId = escapeXml(internetMessageId);
  const request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
    'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
    'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
    'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
      "<soap:Header>" +
        '<t:RequestServerVersion Version="Exchange2013"/>' +
      "</soap:Header>" +
      "<soap:Body>" +
        '<m:FindItem Traversal="Deep">' +
          "<m:ItemShape>" +
            "<t:BaseShape>IdOnly</t:BaseShape>" +
          "</m:ItemShape>" +
          '<m:IndexedPageItemView MaxEntriesReturned="1" Offset="0" BasePoint="Beginning"/>' +
          "<m:Restriction>" +
            "<t:IsEqualTo>" +
              '<t:FieldURI FieldURI="item:InternetMessageId"/>' +
              "<t:FieldURIOrConstant>" +
                '<t:Constant Value="' + escapedId + '"/>' +
              "</t:FieldURIOrConstant>" +
            "</t:IsEqualTo>" +
          "</m:Restriction>" +
          "<m:ParentFolderIds>" +
            '<t:DistinguishedFolderId Id="msgfolderroot"/>' +
          "</m:ParentFolderIds>" +
        "</m:FindItem>" +
      "</soap:Body>" +
    "</soap:Envelope>";

  return makeEwsRequest(request).then((xmlText) => {
    const doc = parseXml(xmlText);
    const message = doc.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "Message")[0];
    if (!message) {
      return "";
    }

    const itemIdNode = message.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "ItemId")[0];
    if (!itemIdNode) {
      return "";
    }

    return String(itemIdNode.getAttribute("Id") || "");
  });
}

function getToRecipientsByItemId(itemId) {
  const escapedItemId = escapeXml(itemId);
  const request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
    'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
    'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
    'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
      "<soap:Header>" +
        '<t:RequestServerVersion Version="Exchange2013"/>' +
      "</soap:Header>" +
      "<soap:Body>" +
        "<m:GetItem>" +
          "<m:ItemShape>" +
            "<t:BaseShape>IdOnly</t:BaseShape>" +
            "<t:AdditionalProperties>" +
              '<t:FieldURI FieldURI="message:ToRecipients"/>' +
            "</t:AdditionalProperties>" +
          "</m:ItemShape>" +
          "<m:ItemIds>" +
            '<t:ItemId Id="' + escapedItemId + '"/>' +
          "</m:ItemIds>" +
        "</m:GetItem>" +
      "</soap:Body>" +
    "</soap:Envelope>";

  return makeEwsRequest(request).then((xmlText) => {
    const doc = parseXml(xmlText);
    const mailboxNodes = doc.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "Mailbox");
    const recipients = [];

    for (let i = 0; i < mailboxNodes.length; i += 1) {
      const emailNode = mailboxNodes[i].getElementsByTagNameNS(
        "http://schemas.microsoft.com/exchange/services/2006/types",
        "EmailAddress"
      )[0];

      if (!emailNode || !emailNode.textContent) {
        continue;
      }

      recipients.push(emailNode.textContent.toLowerCase());
    }

    return recipients;
  });
}

function makeEwsRequest(requestXml) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.makeEwsRequestAsync(requestXml, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        const err = result.error && (result.error.message || result.error.name);
        reject(new Error("EWS request failed: " + String(err || "unknown")));
        return;
      }

      resolve(result.value);
    });
  });
}

function parseXml(xmlText) {
  return new DOMParser().parseFromString(xmlText, "text/xml");
}

function escapeXml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
