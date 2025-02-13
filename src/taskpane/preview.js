Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.addEventListener("DOMContentLoaded", function () {
            const user_info_str = Office.context.roamingSettings.get('user_info');
            if (user_info_str) {
                const user_info = JSON.parse(user_info_str);
                displayUserInfo(user_info);
            }

            document.getElementById("set-button").addEventListener("click", function () {
                // Logic to finalize the details
                alert("Details have been set!");
            });
        });
    }
});

function displayUserInfo(user_info) {
    const userInfoContainer = document.getElementById("user-info-preview");
    userInfoContainer.innerHTML = `
        <div>
            ${user_info.signoff ? `<div>${user_info.signoff}</div>` : ""}
            <div><strong>${user_info.name}</strong></div> ${user_info.pronoun ? `<div> | ${user_info.pronoun}</div>` : ""}
            ${user_info.title ? `<div>${user_info.title}</div>` : ""}
            ${user_info.department ? `<div>${user_info.department}</div>` : ""}
            ${user_info.phone ? `<div>${user_info.phone}</div>` : ""}
            ${user_info.location ? `<div>${user_info.location}</div>` : ""}
        </div>
    `;
}