const form = document.getElementById('connectionForm');
const serverUrl = document.getElementById('serverUrl');
const saveButton = document.getElementById('saveButton');
const message = document.getElementById('message');

window.desktop.getSettings().then(settings => {
  serverUrl.value = settings.serverUrl;
}).catch(error => {
  message.textContent = error.message || 'تعذّر قراءة إعدادات الاتصال.';
});

form.addEventListener('submit', async event => {
  event.preventDefault();
  saveButton.disabled = true;
  message.textContent = '';
  try {
    await window.desktop.saveServerUrl(serverUrl.value);
  } catch (error) {
    message.textContent = error.message || 'تعذّر حفظ عنوان الخادم.';
    saveButton.disabled = false;
  }
});
