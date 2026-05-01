Office.onReady(() => {});

function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  item.from.getAsync((fromResult) => {
    if (fromResult.status !== Office.AsyncResultStatus.Succeeded || !fromResult.value) {
      event.completed({ allowEvent: true });
      return;
    }

    const from = String(fromResult.value.emailAddress || "").toLowerCase();

    item.to.getAsync((toResult) => {
      if (toResult.status !== Office.AsyncResultStatus.Succeeded || !Array.isArray(toResult.value)) {
        event.completed({ allowEvent: true });
        return;
      }

      const recipients = toResult.value
        .map((r) => String(r.emailAddress || "").toLowerCase())
        .filter(Boolean);

      const rules = [
        {
          recipient: "support@pondiot.com",
          validDomain: "@pondiot.com",
          message: "Письмо на support@pondiot.com должно отправляться с адреса домена @pondiot.com"
        },
        {
          recipient: "support@pondmobile.com",
          validDomain: "@pondmobile.com",
          message: "Письмо на support@pondmobile.com должно отправляться с адреса домена @pondmobile.com"
        },
        {
          recipient: "upsupport@pondiot.com",
          exactFrom: "upsupport@pondiot.com",
          message: "Письмо на upsupport@pondiot.com нужно отправлять строго от upsupport@pondiot.com"
        }
      ];

      for (const rule of rules) {
        if (!recipients.includes(rule.recipient)) {
          continue;
        }

        if (rule.validDomain && !from.endsWith(rule.validDomain)) {
          promptUser(event, rule.message);
          return;
        }

        if (rule.exactFrom && from !== rule.exactFrom) {
          promptUser(event, rule.message);
          return;
        }
      }

      event.completed({ allowEvent: true });
    });
  });
}

function promptUser(event, message) {
  event.completed({
    allowEvent: false,
    errorMessage: message
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
