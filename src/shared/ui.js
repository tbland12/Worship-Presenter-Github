const THEME_KEY = 'worship-presenter-theme';

function storedTheme() {
  try {
    const value = window.localStorage.getItem(THEME_KEY);
    return value === 'light' || value === 'dark' ? value : null;
  } catch (error) {
    return null;
  }
}

function applyTheme(theme) {
  const resolved = theme === 'light' ? 'light' : 'dark';
  document.documentElement.dataset.theme = resolved;
  return resolved;
}

export function initializeAppearance(button) {
  let theme = applyTheme(storedTheme() || 'dark');
  const updateButton = () => {
    if (!button) return;
    const next = theme === 'dark' ? 'light' : 'dark';
    button.setAttribute('aria-label', `Switch to ${next} appearance`);
    button.setAttribute('title', `Switch to ${next} appearance`);
    button.dataset.theme = theme;
  };
  updateButton();
  if (button) {
    button.addEventListener('click', () => {
      theme = applyTheme(theme === 'dark' ? 'light' : 'dark');
      try {
        window.localStorage.setItem(THEME_KEY, theme);
      } catch (error) {
        // Appearance still applies for the current window when storage is unavailable.
      }
      updateButton();
    });
  }
  return theme;
}

export function createNotifier(region) {
  const target = region || document.getElementById('toast-region');
  return function notify(options = {}) {
    if (!target) return null;
    const type = ['success', 'warning', 'error'].includes(options.type) ? options.type : 'info';
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.dataset.type = type;
    toast.setAttribute('role', type === 'error' ? 'alert' : 'status');

    const accent = document.createElement('span');
    accent.className = 'toast-accent';
    accent.setAttribute('aria-hidden', 'true');

    const copy = document.createElement('div');
    copy.className = 'toast-copy';
    const title = document.createElement('div');
    title.className = 'toast-title';
    title.textContent = options.title || (type === 'error' ? 'Something went wrong' : 'Notice');
    const message = document.createElement('div');
    message.className = 'toast-message';
    message.textContent = options.message || '';
    copy.append(title, message);

    const dismiss = document.createElement('button');
    dismiss.type = 'button';
    dismiss.className = 'toast-dismiss';
    dismiss.setAttribute('aria-label', 'Dismiss notification');
    dismiss.textContent = '×';
    const remove = () => toast.remove();
    dismiss.addEventListener('click', remove);
    toast.append(accent, copy, dismiss);
    target.appendChild(toast);

    const timeout = options.timeout ?? (type === 'error' ? 0 : 5000);
    if (timeout > 0) window.setTimeout(remove, timeout);
    return toast;
  };
}
