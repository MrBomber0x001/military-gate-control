const stateElement = document.getElementById('state');
const stateText = document.getElementById('stateText');
const errorElement = document.getElementById('error');
const addressList = document.getElementById('addresses');
const portForm = document.getElementById('portForm');
const portInput = document.getElementById('portInput');
const retryButton = document.getElementById('retryButton');
const params = new URLSearchParams(location.search);
let port = params.get('port') || '3000';
let addresses = [];

try {
  addresses = JSON.parse(params.get('addresses') || '[]');
} catch (error) {
  console.error('Unable to read server addresses:', error);
}

function renderAddresses() {
  const options = [`http://localhost:${port}`, ...addresses];
  addressList.replaceChildren(...[...new Set(options)].map(address => {
    const row = document.createElement('div');
    row.className = 'address';
    const code = document.createElement('code');
    code.textContent = address;
    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = 'فتح';
    button.addEventListener('click', () => {
      window.gateServer.openAddress(address).catch(error => {
        errorElement.textContent = error.message || 'تعذّر فتح العنوان.';
      });
    });
    row.append(code, button);
    return row;
  }));
}

function renderStatus({ status: currentStatus, message = '', port: serverPort, addresses: serverAddresses }) {
  const running = currentStatus === 'running';
  if (serverPort) {
    port = String(serverPort);
    portInput.value = port;
  }
  if (Array.isArray(serverAddresses)) addresses = serverAddresses;
  stateElement.classList.toggle('running', running);
  stateText.textContent = running
    ? 'الخادم يعمل وجاهز للاتصال'
    : currentStatus === 'starting' ? 'جارٍ تشغيل الخادم…' : 'الخادم متوقف أو يتعذّر تشغيله';
  errorElement.textContent = message;
  portForm.hidden = currentStatus !== 'port-in-use';
  renderAddresses();
}

portInput.value = port;
portForm.addEventListener('submit', async event => {
  event.preventDefault();
  if (!portForm.reportValidity()) return;
  retryButton.disabled = true;
  errorElement.textContent = '';
  try {
    await window.gateServer.retryPort(portInput.value);
  } catch (error) {
    errorElement.textContent = error.message || 'تعذّر تشغيل الخادم على المنفذ المحدد.';
  } finally {
    retryButton.disabled = false;
  }
});

renderAddresses();
window.gateServer.getStatus().then(renderStatus).catch(error => {
  renderStatus({ status: 'stopped', message: error.message });
});
window.gateServer.onStatus(renderStatus);
