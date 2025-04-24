window.onload = async function() {
    const settings = await eel.get_settings()();
    document.getElementById("lang").value = settings.lang;
    document.getElementById("signature").value = settings.signature;
};

async function saveSettings() {
    const lang = document.getElementById("lang").value;
    const signature = document.getElementById("signature").value;

    const result = await eel.update_settings(lang,signature)();
    document.getElementById("saveStatus").innerText = result;

    setTimeout(() => {
        window.location.href = "index.html";
    }, 1000);
}