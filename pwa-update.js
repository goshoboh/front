(() => {
  'use strict';

  const currentVersion = document.querySelector('meta[name="app-version"]')?.content;
  const notice = document.getElementById('appUpdateNotice');
  const message = document.getElementById('appUpdateMessage');
  const updateButton = document.getElementById('appUpdateButton');
  const versionUrl = new URL('version.json', document.baseURI);
  let registration = null;
  let checking = false;

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', async () => {
      try {
        registration = await navigator.serviceWorker.register('sw.js', {
          updateViaCache: 'none'
        });
      } catch (error) {
        console.warn('Service Worker の登録に失敗しました:', error);
      }
    });
  }

  // 取得した最新版を現在の版にせず、読み込んだ HTML の版と比較する。
  async function checkForUpdate() {
    if (!currentVersion || checking || document.hidden || !navigator.onLine) return;
    checking = true;
    try {
      if (registration) {
        registration.update().catch(error => {
          console.warn('Service Worker の更新確認に失敗しました:', error);
        });
      }
      const url = new URL(versionUrl);
      url.searchParams.set('_', Date.now().toString());
      const response = await fetch(url, { cache: 'no-store' });
      if (!response.ok) return;
      const release = await response.json();
      if (typeof release.version === 'string' && release.version !== currentVersion) {
        notice.hidden = false;
      }
    } catch (error) {
      // 一時的な通信失敗で、台帳の利用を妨げない。
      console.warn('アプリの更新確認に失敗しました:', error);
    } finally {
      checking = false;
    }
  }

  updateButton.addEventListener('click', () => {
    const editors = document.querySelectorAll(
      '.pickwindow, #guestModal, #roomMoveModal, #carPicker'
    );
    if (Array.from(editors).some(editor => getComputedStyle(editor).display !== 'none')) {
      message.textContent = '入力画面を閉じ、保存が終わってから更新してください。';
      return;
    }
    window.location.reload();
  });

  document.addEventListener('visibilitychange', checkForUpdate);
  window.addEventListener('focus', checkForUpdate);
  window.addEventListener('online', checkForUpdate);
  setInterval(checkForUpdate, 5 * 60 * 1000);
  checkForUpdate();
})();
