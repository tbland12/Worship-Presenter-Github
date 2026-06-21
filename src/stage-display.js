const elements = {
  layout: document.getElementById('stage-layout'),
  section: document.getElementById('stage-section'),
  title: document.getElementById('stage-title'),
  clock: document.getElementById('stage-clock'),
  currentLabel: document.getElementById('current-label'),
  currentText: document.getElementById('current-text'),
  notesText: document.getElementById('notes-text'),
  nextLabel: document.getElementById('next-label'),
  nextText: document.getElementById('next-text'),
  panicBanner: document.getElementById('panic-banner')
};

function updateClock() {
  const now = new Date();
  elements.clock.dateTime = now.toISOString();
  elements.clock.textContent = new Intl.DateTimeFormat(undefined, {
    hour: 'numeric',
    minute: '2-digit',
    second: '2-digit'
  }).format(now);
}

function renderState(state = {}) {
  const active = state.active === true;
  elements.layout.dataset.active = active ? 'true' : 'false';
  elements.layout.dataset.panic = state.panic ? 'true' : 'false';
  elements.section.textContent = active ? state.section || 'Live' : 'Stage display';
  elements.title.textContent = active ? state.itemTitle || 'Live content' : 'Waiting for live content';
  elements.currentLabel.textContent = active ? state.currentLabel || 'Current' : 'Current';
  elements.currentText.textContent = active ? state.currentText || 'Media slide' : 'Nothing is live';
  elements.notesText.textContent = active && state.currentNotes ? state.currentNotes : 'No notes for this slide';
  elements.nextLabel.textContent = active && state.nextLabel ? `· ${state.nextLabel}` : '';
  elements.nextText.textContent = active && state.nextText ? state.nextText : 'End of section';
  elements.panicBanner.hidden = !state.panic;
}

window.api?.onStageState?.(renderState);
document.addEventListener('keydown', (event) => {
  if (event.key !== 'Escape') return;
  event.preventDefault();
  window.api?.hideStage?.();
});

updateClock();
window.setInterval(updateClock, 1000);
