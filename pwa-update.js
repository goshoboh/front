(() => {
  'use strict';

  if (!('serviceWorker' in navigator)) return;

  // アプリ本体の更新はページ再読み込み時に取り込む。
  // Worker の更新でも、利用中の画面を強制再読み込みしない。
  async function registerWorker() {
    try {
      await navigator.serviceWorker.register('sw.js', {
        updateViaCache: 'none'
      });
    } catch (error) {
      console.warn('Service Worker の登録・更新に失敗しました:', error);
    }
  }

  window.addEventListener('load', registerWorker);
  window.addEventListener('online', registerWorker);
})();
