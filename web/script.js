async function sendMail() {
    const issue = document.getElementById("issueInput").value;
    const response = await eel.send_email(issue)();
    document.getElementById("status").innerText = response;
}