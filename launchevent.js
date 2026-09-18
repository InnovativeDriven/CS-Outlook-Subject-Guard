const SUBJECT_PATTERN = /^\[([^|\]]+)\|([^|\]]+)\](.*)$/;

const ERROR_TEXT =
  "This subject has no client/matter tag. " +
  "Expected format: [Acme Corp | Smith Arbitration] Production volume 3 questions";

function subjectIsCompliant(subject) {
  const m = SUBJECT_PATTERN.exec(subject || "");
  if (!m) return false;

  const client = m[1].trim();
  const matter = m[2].trim();
  const rest   = m[3].trim();

  return client !== "" && matter !== "" && rest !== "";
}

      const subject = result.value || "";

      if (SUBJECT_PATTERN.test(subject)) {
        evt.completed({ allowEvent: true });
      } else {
        evt.completed({
          allowEvent: false,
          errorMessage: ERROR_TEXT
        });
      }
    }
  );
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);