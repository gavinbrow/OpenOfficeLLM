// The service worker. Deliberately almost empty.
//
// MV3 service workers are killed and restarted at the browser's discretion, so
// anything stateful put here is state that vanishes without warning. The agent
// loop, the conversation, the adapter and the host connection all live in the
// side panel, which stays alive as long as it is open. This file exists for the
// two jobs only the worker can do.

// 1. Make the toolbar icon open the panel. Without this the action click is a
//    no-op and the panel is only reachable from the puzzle-piece menu.
chrome.runtime.onInstalled.addListener(() => {
  void chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch(() => {
    // Older Chromium builds have no sidePanel.setPanelBehavior. The panel is
    // still reachable manually, so this is a downgrade, not a failure.
  })
})

// 2. Keep the panel enabled across tabs. The panel follows the window, and a
//    per-tab disable would make it disappear on a tab we cannot read — which is
//    wrong: chat still works there, only page tools are unavailable, and the
//    panel says so itself.
chrome.tabs.onActivated.addListener(({ tabId }) => {
  void chrome.sidePanel.setOptions({ tabId, enabled: true }).catch(() => {})
})
