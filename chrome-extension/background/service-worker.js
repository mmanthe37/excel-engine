/* Excel Engine — Background Service Worker */

chrome.runtime.onInstalled.addListener(() => {
  console.log('Excel Engine extension installed');
  chrome.storage.sync.set({
    serverURL: 'http://localhost:8080',
    apiKey: '',
  });
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === 'health-check') {
    chrome.storage.sync.get({ serverURL: 'http://localhost:8080' }, async (items) => {
      try {
        const resp = await fetch(`${items.serverURL}/health`);
        const data = await resp.json();
        sendResponse({ ok: true, data });
      } catch (err) {
        sendResponse({ ok: false, error: err.message });
      }
    });
    return true;
  }
});
