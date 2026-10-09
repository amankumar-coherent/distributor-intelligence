const DEFAULT_URL = "http://localhost:15551";

document.addEventListener("DOMContentLoaded", async () => {
  const input = document.getElementById("serverUrl");
  // Server URL lives in storage.LOCAL, not sync. chrome.storage.sync
  // replicates across every profile signed into the same Google account, so
  // pointing one profile at :15561 silently rewrote the others and several
  // "isolated" browsers ended up polling the same server — scraping tabs
  // then opened in whichever profile happened to poll first.
  // storage.local is per-profile and never syncs.
  const local = await chrome.storage.local.get({ serverUrl: "" });
  let serverUrl = local.serverUrl;
  if (!serverUrl) {
    // One-time migration for profiles configured before this change.
    const legacy = await chrome.storage.sync.get({ serverUrl: DEFAULT_URL });
    serverUrl = legacy.serverUrl || DEFAULT_URL;
  }
  input.value = serverUrl;

  document.getElementById("save").addEventListener("click", async () => {
    const url = input.value.trim().replace(/\/+$/, "") || DEFAULT_URL;
    input.value = url;
    await chrome.storage.local.set({ serverUrl: url });
    const status = document.getElementById("status");
    status.classList.add("visible");
    setTimeout(() => status.classList.remove("visible"), 1500);
  });
});
