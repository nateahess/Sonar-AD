// DOM Elements
const domainDisplay = document.getElementById('domainDisplay');
const generateBtn = document.getElementById('generateBtn');
const consoleOutput = document.getElementById('console');
const clearConsoleBtn = document.getElementById('clearConsoleBtn');
const statusMessage = document.getElementById('statusMessage');

let isGenerating = false;
let currentReportPath = null;

// Initialize app
async function init() {
  await loadDomain();
  setupEventListeners();
  setupScriptOutputListener();
}

// Load AD Domain
async function loadDomain() {
  try {
    const result = await window.electronAPI.getDomain();

    if (result.success) {
      domainDisplay.innerHTML = `
        <svg class="domain-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M21 12a9 9 0 01-9 9m9-9a9 9 0 00-9-9m9 9H3m9 9a9 9 0 01-9-9m9 9c1.657 0 3-4.03 3-9s-1.343-9-3-9m0 18c-1.657 0-3-4.03-3-9s1.343-9 3-9m-9 9a9 9 0 019-9"></path>
        </svg>
        <span>${result.domain}</span>
      `;
      domainDisplay.classList.add('loaded');
      generateBtn.disabled = false;
      logToConsole(`Connected to domain: ${result.domain}`, 'success');
    } else {
      domainDisplay.innerHTML = `
        <svg class="domain-icon error" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <circle cx="12" cy="12" r="10"></circle>
          <line x1="12" y1="8" x2="12" y2="12"></line>
          <line x1="12" y1="16" x2="12.01" y2="16"></line>
        </svg>
        <span>Not Connected</span>
      `;
      domainDisplay.classList.add('error');
      logToConsole(`Error: ${result.error || 'Unable to retrieve domain'}`, 'error');
      logToConsole(result.message || '', 'warning');

      showStatus('Unable to connect to domain. Make sure you are on a domain network.', 'error');
    }
  } catch (error) {
    domainDisplay.innerHTML = `
      <svg class="domain-icon error" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <circle cx="12" cy="12" r="10"></circle>
        <line x1="12" y1="8" x2="12" y2="12"></line>
        <line x1="12" y1="16" x2="12.01" y2="16"></line>
      </svg>
      <span>Error</span>
    `;
    domainDisplay.classList.add('error');
    logToConsole(`Fatal error: ${error.message}`, 'error');
  }
}

// Setup Event Listeners
function setupEventListeners() {
  generateBtn.addEventListener('click', handleGenerateReport);
  clearConsoleBtn.addEventListener('click', clearConsole);
}

// Setup Script Output Listener
function setupScriptOutputListener() {
  window.electronAPI.onScriptOutput((data) => {
    logToConsole(data, 'info');
  });
}

// Handle Generate Report
async function handleGenerateReport() {
  if (isGenerating) return;

  isGenerating = true;
  generateBtn.disabled = true;
  generateBtn.classList.add('loading');

  const btnText = generateBtn.querySelector('.btn-text');
  const originalText = btnText.textContent;
  btnText.textContent = 'Generating...';

  logToConsole('\n=== Starting Report Generation ===', 'info');
  clearStatus();

  try {
    const result = await window.electronAPI.generateReport();

    if (result.success) {
      logToConsole('\n=== Report Generation Complete ===', 'success');
      logToConsole(`Report saved to: ${result.reportPath}`, 'success');

      currentReportPath = result.reportPath;

      showStatus('Report generated successfully! Opening report...', 'success');

      // Auto-open the report
      setTimeout(async () => {
        await window.electronAPI.openReport(result.reportPath);
        showStatus('Report opened in browser', 'success');
      }, 500);

    } else {
      logToConsole('\n=== Report Generation Failed ===', 'error');
      logToConsole(`Error: ${result.error}`, 'error');
      if (result.message) {
        logToConsole(result.message, 'warning');
      }

      showStatus(`Failed to generate report: ${result.error}`, 'error');
    }
  } catch (error) {
    logToConsole(`\nFatal error: ${error.message}`, 'error');
    showStatus('An unexpected error occurred', 'error');
  } finally {
    isGenerating = false;
    generateBtn.disabled = false;
    generateBtn.classList.remove('loading');
    btnText.textContent = originalText;
  }
}

// Log to Console
function logToConsole(message, type = 'info') {
  const logEntry = document.createElement('div');
  logEntry.className = `console-entry console-${type}`;

  const timestamp = new Date().toLocaleTimeString();
  const timestampSpan = document.createElement('span');
  timestampSpan.className = 'console-timestamp';
  timestampSpan.textContent = `[${timestamp}] `;

  const messageSpan = document.createElement('span');
  messageSpan.textContent = message;

  logEntry.appendChild(timestampSpan);
  logEntry.appendChild(messageSpan);

  consoleOutput.appendChild(logEntry);
  consoleOutput.scrollTop = consoleOutput.scrollHeight;
}

// Clear Console
function clearConsole() {
  consoleOutput.innerHTML = '';
  logToConsole('Console cleared', 'info');
}

// Show Status Message
function showStatus(message, type = 'info') {
  statusMessage.textContent = message;
  statusMessage.className = `status-message status-${type}`;
  statusMessage.classList.remove('hidden');

  // Auto-hide after 5 seconds for success messages
  if (type === 'success') {
    setTimeout(() => {
      statusMessage.classList.add('hidden');
    }, 5000);
  }
}

// Clear Status Message
function clearStatus() {
  statusMessage.classList.add('hidden');
}

// Initialize on load
init();
