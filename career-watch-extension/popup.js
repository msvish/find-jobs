const form = document.getElementById("form");
const company = document.getElementById("company");
const url = document.getElementById("url");
const save = document.getElementById("save");
const status = document.getElementById("status");

function show(msg, kind = "") {
  status.textContent = msg;
  status.className = kind;
}

// Prefill the careers page with the tab you're on
chrome.tabs.query({ active: true, currentWindow: true }, ([tab]) => {
  if (tab?.url?.startsWith("http")) url.value = tab.url;
  company.focus();
});

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  save.disabled = true;
  show("Adding...");
  try {
    const res = await fetch(WEB_APP_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain" },
      body: JSON.stringify({ token: TOKEN, company: company.value.trim(), url: url.value.trim() }),
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || "The sheet rejected the row.");
    show(`Added ${company.value.trim()} to row ${data.row}.`, "ok");
    company.value = "";
  } catch (err) {
    show(err.message.includes("JSON") ? "Couldn't reach the sheet. Check WEB_APP_URL in config.js." : err.message, "err");
  } finally {
    save.disabled = false;
  }
});
