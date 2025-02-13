Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    initializeFabricUI();

    checkUserInfo();
  }
});

function initializeFabricUI() {
  // You might need to manually initialize certain Fabric components
  // For example, dropdowns:
  var elements = document.querySelectorAll('.ms-Dropdown');
  for (var i = 0; i < elements.length; i++) {
      new fabric['Dropdown'](elements[i]);
  }
}

function checkUserInfo() {
  const user_info_str = Office.context.roamingSettings.get('user_info');
  if (user_info_str) {
    const user_info = JSON.parse(user_info_str);
    loadSignature(user_info);
  } else {
    // TODO: notification to open task pane
  }
}

function saveUserInfo(event) {
  event.preventDefault();
  console.log("Saving user info...");

  const name = document.getElementById("name").value;
  const title = document.getElementById("title").value;

  if (!name || !title) {
    alert("Please fill in all required fields.");
    return;
}

  const user_info = {
    name: document.getElementById("name").value.trim(),
    title: document.getElementById("title").value.trim(),
    department: document.getElementById("department").value.trim(),
    phone: document.getElementById("phone").value.trim(),
    location: document.getElementById("location").value.trim(),
    pronoun: document.getElementById("pronoun").value.trim(),
    signoff: document.getElementById("signoff").value.trim()
  };

  console.log("User info:", user_info);
  Office.context.roamingSettings.set('user_info', JSON.stringify(user_info));
  Office.context.roamingSettings.saveAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log("Failed to save user info. Please try again.");
    } else {
      console.log("User info saved successfully!");
      window.location.href = "preview.html";
    }
  });
}

function loadSignature(user_info) {
  Office.context.mailbox.item.getComposeTypeAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const composeType = asyncResult.value;
      console.log("Compose type:", composeType);

      if (composeType === Office.MailboxEnums.ComposeType.NewMail) {
        addSignatureToBody(user_info, "newMail");
      } else if (composeType === Office.MailboxEnums.ComposeType.Reply) {
        addSignatureToBody(user_info, "reply");
      } else if (composeType === Office.MailboxEnums.ComposeType.Forward) {
        addSignatureToBody(user_info, "reply");
      }
    } else {
      console.error("Failed to get compose type:", asyncResult.error.message);
    }
  });
}

function addSignatureToBody(user_info, composeType) {
  let signature = generateSignature(user_info, composeType);
  Office.context.mailbox.item.body.setAsync(
    signature,
    { coercionType: Office.CoercionType.Html },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to set signature:", asyncResult.error.message);
      } else {
        console.log("Signature set successfully.");
      }
    }
  );
}

function generateSignature(user_info, composeType) {
  if (composeType === "newMail") {
    let signature = `
      <div>
        ${user_info.signoff ? `<div>${user_info.signoff}</div>` : ""}
        <div><strong>${user_info.name}</strong></div> ${user_info.pronoun ? `<div> | ${user_info.pronoun}</div>` : ""}
        ${user_info.title ? `<div>${user_info.title}</div>` : ""}
        ${user_info.department ? `<div>${user_info.department}</div>` : ""}
        ${user_info.phone ? `<div>${user_info.phone}</div>` : ""}
        ${user_info.location ? `<div>${user_info.location}</div>` : ""}
      </div>
    `;
    return signature;
  } else if (composeType === "reply") {
    let signature = `
      <div>
        ${user_info.signoff ? `<div>${user_info.signoff}</div>` : ""}
        <div><strong>${user_info.name}</strong></div>
        ${user_info.title ? `<div>${user_info.title}</div>` : ""}
        ${user_info.department ? `<div>${user_info.department}</div>` : ""}
        ${user_info.phone ? `<div>${user_info.phone}</div>` : ""}
        ${user_info.location ? `<div>${user_info.location}</div>` : ""}
        ${user_info.pronoun ? `<div>${user_info.pronoun}</div>` : ""}
      </div>
    `;
    return signature;
  }
}