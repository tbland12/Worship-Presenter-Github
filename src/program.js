import { StageRenderer } from './shared/stage.js';
import { ensureApiBridge } from './shared/bridge.js';

const stageRoot = document.getElementById('program-stage');
const stage = new StageRenderer(stageRoot);
let currentState = null;

ensureApiBridge();

window.api.onProgramState((state) => {
  currentState = state;
  stage.render(state);
});

stage.onVideoEnded = () => {
  if (!currentState || currentState.section !== 'timer') {
    return;
  }
  if (window.api && window.api.sendProgramEvent) {
    window.api.sendProgramEvent({ type: 'video-ended' });
  }
};

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') {
    event.preventDefault();
    window.api.hideProgram();
  }
});
