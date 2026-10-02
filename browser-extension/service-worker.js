const ENDPOINT = 'http://127.0.0.1:8765/quick-capture/';

chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.create({
    id: 'procurement-quick-capture',
    title: 'Procurement Control',
    contexts: ['selection']
  });
  chrome.contextMenus.create({
    id: 'procurement-task',
    parentId: 'procurement-quick-capture',
    title: 'Создать задачу',
    contexts: ['selection']
  });
  chrome.contextMenus.create({
    id: 'procurement-reminder',
    parentId: 'procurement-quick-capture',
    title: 'Создать напоминание',
    contexts: ['selection']
  });
});

chrome.contextMenus.onClicked.addListener(async (info) => {
  const kind = info.menuItemId === 'procurement-task' ? 'task' :
    info.menuItemId === 'procurement-reminder' ? 'reminder' : null;
  if (!kind || !info.selectionText) return;

  try {
    const response = await fetch(ENDPOINT, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ kind, text: info.selectionText })
    });
    const result = await response.json();
    if (!response.ok || !result.ok) throw new Error(result.error || 'Агент отклонил запрос.');
  } catch (error) {
    chrome.notifications.create({
      type: 'basic',
      iconUrl: 'icon.png',
      title: 'Procurement Control',
      message: 'Не удалось открыть форму. Запустите Procurement Control.'
    });
  }
});
