// Auto Greet - event handler for Outlook event-based activation.
// Loaded directly by classic Outlook on Windows, and via commands.html
// by Outlook on the web, new Outlook on Windows, and Outlook on Mac.

function getFirstName(name, email) {
  if (name && name.trim().length > 0) {
    return capitalise(name.trim().split(/\s+/)[0]);
  }
  if (email && email.indexOf("@") !== -1) {
    return capitalise(email.split("@")[0].split(/[._\-+]/)[0]);
  }
  return null;
}

function capitalise(word) {
  if (!word) return word;
  return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
}

// Fires automatically when a compose window opens.
function autoGreet(event) {
  var item = Office.context.mailbox.item;

  // Only act on a Reply / Reply All -- skip brand-new emails and forwards.
  item.getComposeTypeAsync(function (typeResult) {
    if (typeResult.status !== Office.AsyncResultStatus.Succeeded ||
        typeResult.value.composeType !== "reply") {
      event.completed();
      return;
    }

    // On a reply, Outlook has already put the original sender in the To field.
    item.to.getAsync(function (toResult) {
      var firstName = null;
      if (toResult.status === Office.AsyncResultStatus.Succeeded &&
          toResult.value && toResult.value.length > 0) {
        firstName = getFirstName(toResult.value[0].displayName,
                                 toResult.value[0].emailAddress);
      }

      if (!firstName) {
        event.completed();
        return;
      }

      var greeting = "<p>Hi " + firstName + ",</p><p><br></p>";

      // Read the existing reply body (quoted thread + signature) and keep it.
      item.body.getAsync(Office.CoercionType.Html, function (bodyResult) {
        if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
          event.completed();
          return;
        }

        var existing = bodyResult.value || "";

        // Don't add the greeting twice.
        if (existing.indexOf("Hi " + firstName + ",") !== -1) {
          event.completed();
          return;
        }

        item.body.setAsync(
          greeting + existing,
          { coercionType: Office.CoercionType.Html },
          function () { event.completed(); }
        );
      });
    });
  });
}

// Connect the function above to the trigger named in the manifest.
Office.actions.associate("autoGreet", autoGreet);
