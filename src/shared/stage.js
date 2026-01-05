function buildTextShadow(style, scale = 1) {
  const shadows = [];
  const stroke = Number(style.strokeWidthPx || 0) * scale;
  const strokeColor = style.strokeColor || '#000000';
  if (stroke > 0) {
    const offsets = [
      [stroke, 0],
      [-stroke, 0],
      [0, stroke],
      [0, -stroke],
      [stroke, stroke],
      [stroke, -stroke],
      [-stroke, stroke],
      [-stroke, -stroke]
    ];
    offsets.forEach(([x, y]) => {
      shadows.push(`${x}px ${y}px 0 ${strokeColor}`);
    });
  }
  if (style.shadow && style.shadow.blur >= 0) {
    const dx = (style.shadow.dx || 0) * scale;
    const dy = (style.shadow.dy || 0) * scale;
    const blur = (style.shadow.blur || 0) * scale;
    const color = style.shadow.color || '#00000088';
    shadows.push(`${dx}px ${dy}px ${blur}px ${color}`);
  }
  return shadows.join(', ');
}

function buildBackgroundKey(background) {
  if (!background || !background.path) {
    return 'none';
  }
  return `${background.type}:${background.path}`;
}

function resolveTextStyle(theme, key) {
  const basePx = theme.baseFontPx || 70;
  let defaultFontPx = basePx;
  if (key === 'title') {
    defaultFontPx = 70;
  } else if (key === 'footer') {
    defaultFontPx = 25;
  }

  const fallbackShadow = theme.shadow || { dx: 2, dy: 2, blur: 6, color: '#000000' };
  const style = (theme.textStyles && theme.textStyles[key]) || {};
  const shadow = style.shadow || {};

  return {
    fontFamily: style.fontFamily || theme.fontFamily || 'Segoe UI',
    fontPx: style.fontPx || defaultFontPx,
    color: style.color || theme.textColor || '#FFFFFF',
    strokeWidthPx: style.strokeWidthPx ?? theme.strokeWidthPx ?? 1,
    strokeColor: style.strokeColor || theme.strokeColor || '#000000',
    shadow: {
      dx: shadow.dx ?? fallbackShadow.dx ?? 0,
      dy: shadow.dy ?? fallbackShadow.dy ?? 0,
      blur: shadow.blur ?? fallbackShadow.blur ?? 0,
      color: shadow.color || fallbackShadow.color || '#000000'
    }
  };
}

export class StageRenderer {
  constructor(root, options = {}) {
    this.root = root;
    this.scale = options.scale ?? 1;
    this.root.classList.add('stage-root');

    this.bgCrossfadeMs = 700;
    this.textFadeMs = 200;
    this.safeMarginsPct = 5;

    this.activeTrack = 'A';
    this.currentBackgroundKey = null;
    this.currentSlideKey = null;
    this.isPanic = false;
    this.currentTheme = null;

    this.bgToken = 0;
    this.textToken = 0;

    this.bgContainer = document.createElement('div');
    this.bgContainer.className = 'stage-bg';

    this.trackA = this.createTrack('A');
    this.trackB = this.createTrack('B');

    this.bgContainer.appendChild(this.trackA.container);
    this.bgContainer.appendChild(this.trackB.container);

    this.dimOverlay = document.createElement('div');
    this.dimOverlay.className = 'stage-dim';

    this.textLayer = document.createElement('div');
    this.textLayer.className = 'stage-text pos-center';

    this.textBox = document.createElement('div');
    this.textBox.className = 'stage-text-box';

    this.titleEl = document.createElement('div');
    this.titleEl.className = 'stage-title';

    this.lyricsEl = document.createElement('div');
    this.lyricsEl.className = 'stage-lyrics';

    this.footerEl = document.createElement('div');
    this.footerEl.className = 'stage-footer';

    this.textBox.appendChild(this.titleEl);
    this.textBox.appendChild(this.lyricsEl);
    this.textBox.appendChild(this.footerEl);
    this.textLayer.appendChild(this.textBox);

    this.root.appendChild(this.bgContainer);
    this.root.appendChild(this.dimOverlay);
    this.root.appendChild(this.textLayer);

    this.onVideoEnded = null;
  }

  clear() {
    this.currentBackgroundKey = null;
    this.currentSlideKey = null;
    this.setBackgroundImmediate(null);
    this.applyText({
      title: '',
      lyrics: '',
      footer: '',
      showTitle: false,
      showLyrics: false,
      showFooter: false
    });
    this.setPanic(false);
  }

  createTrack(label) {
    const container = document.createElement('div');
    container.className = 'stage-bg-track';
    container.dataset.track = label;

    const video = document.createElement('video');
    video.className = 'stage-bg-video';
    video.muted = true;
    video.loop = true;
    video.playsInline = true;
    video.preload = 'auto';
    video.autoplay = true;
    video.defaultMuted = true;

    const img = document.createElement('img');
    img.className = 'stage-bg-image';
    img.decoding = 'async';
    img.loading = 'eager';

    const track = { container, video, img, type: null, src: null };

    container.appendChild(video);
    container.appendChild(img);

    video.addEventListener('ended', () => {
      if (this.onVideoEnded && this.getActiveTrack() === track) {
        this.onVideoEnded();
      }
    });
    return track;
  }

  getActiveTrack() {
    return this.activeTrack === 'A' ? this.trackA : this.trackB;
  }

  setSettings(settings = {}) {
    if (settings.textFadeMs != null) {
      this.textFadeMs = settings.textFadeMs;
    }
    if (settings.bgCrossfadeMs != null) {
      this.bgCrossfadeMs = settings.bgCrossfadeMs;
    }
    if (settings.safeMarginsPct != null) {
      this.safeMarginsPct = settings.safeMarginsPct;
      this.root.style.setProperty('--safe-margins-pct', String(settings.safeMarginsPct));
    }
  }

  setTheme(theme) {
    if (!theme) {
      return;
    }
    this.currentTheme = theme;
    const scale = this.scale || 1;
    const titleStyle = resolveTextStyle(theme, 'title');
    const lyricsStyle = resolveTextStyle(theme, 'lyrics');
    const footerStyle = resolveTextStyle(theme, 'footer');

    this.root.style.setProperty('--title-font-family', titleStyle.fontFamily);
    this.root.style.setProperty('--title-font-size', `${titleStyle.fontPx * scale}px`);
    this.root.style.setProperty('--title-text-color', titleStyle.color);
    this.root.style.setProperty('--title-text-shadow', buildTextShadow(titleStyle, scale));

    this.root.style.setProperty('--lyrics-font-family', lyricsStyle.fontFamily);
    this.root.style.setProperty('--lyrics-font-size', `${lyricsStyle.fontPx * scale}px`);
    this.root.style.setProperty('--lyrics-text-color', lyricsStyle.color);
    this.root.style.setProperty('--lyrics-text-shadow', buildTextShadow(lyricsStyle, scale));

    this.root.style.setProperty('--footer-font-family', footerStyle.fontFamily);
    this.root.style.setProperty('--footer-font-size', `${footerStyle.fontPx * scale}px`);
    this.root.style.setProperty('--footer-text-color', footerStyle.color);
    this.root.style.setProperty('--footer-text-shadow', buildTextShadow(footerStyle, scale));

    this.dimOverlay.style.opacity = String(theme.dimOpacity != null ? theme.dimOpacity : 0.25);

    if (theme.position === 'lower-third') {
      this.textLayer.classList.remove('pos-center');
      this.textLayer.classList.add('pos-lower');
    } else {
      this.textLayer.classList.remove('pos-lower');
      this.textLayer.classList.add('pos-center');
    }

    this.minFontPx = (theme.minFontPx || 36) * scale;
    this.baseFontPx = (lyricsStyle.fontPx || theme.baseFontPx || 64) * scale;
  }

  setScale(scale = 1) {
    this.scale = scale;
    if (this.currentTheme) {
      this.setTheme(this.currentTheme);
    }
  }

  async render(state) {
    if (!state) {
      return;
    }
    if (state.settings) {
      this.setSettings(state.settings);
    }
    if (state.theme) {
      this.setTheme(state.theme);
    }

    const newBackgroundKey = state.backgroundKey || buildBackgroundKey(state.background);
    if (this.currentBackgroundKey === null) {
      this.setBackgroundImmediate(state.background);
      this.currentBackgroundKey = newBackgroundKey;
    } else if (newBackgroundKey !== this.currentBackgroundKey) {
      this.crossfadeBackground(state.background);
      this.currentBackgroundKey = newBackgroundKey;
    }

    if (state.textImmediate) {
      if (state.slideKey) {
        this.currentSlideKey = state.slideKey;
      }
      this.applyText(state.text);
      this.applyAutoFit();
      this.setPanic(state.panic);
      return;
    }

    if (state.slideKey && state.slideKey !== this.currentSlideKey) {
      this.currentSlideKey = state.slideKey;
      await this.transitionText(state.text, state.panic);
    } else {
      this.applyText(state.text);
      this.applyAutoFit();
      this.setPanic(state.panic);
    }
  }

  setBackgroundImmediate(background) {
    if (!background || !background.path) {
      this.setTrackContent(this.trackA, null);
      this.setTrackContent(this.trackB, null);
      this.trackA.container.style.opacity = '0';
      this.trackB.container.style.opacity = '0';
      return;
    }
    const track = this.activeTrack === 'A' ? this.trackA : this.trackB;
    this.setTrackContent(track, background, { forceReload: true });
    track.container.style.opacity = '1';
    const other = track === this.trackA ? this.trackB : this.trackA;
    other.container.style.opacity = '0';
  }

  crossfadeBackground(background) {
    const token = ++this.bgToken;
    const incoming = this.activeTrack === 'A' ? this.trackB : this.trackA;
    const outgoing = incoming === this.trackA ? this.trackB : this.trackA;

    this.setTrackContent(incoming, background, { forceReload: true });
    incoming.container.style.transition = `opacity ${this.bgCrossfadeMs}ms ease`;
    outgoing.container.style.transition = `opacity ${this.bgCrossfadeMs}ms ease`;
    incoming.container.style.opacity = '0';
    outgoing.container.style.opacity = '1';

    requestAnimationFrame(() => {
      if (token !== this.bgToken) {
        return;
      }
      incoming.container.style.opacity = '1';
      outgoing.container.style.opacity = '0';
    });

    window.setTimeout(() => {
      if (token !== this.bgToken) {
        return;
      }
      this.stopTrack(outgoing);
      this.activeTrack = incoming === this.trackA ? 'A' : 'B';
    }, this.bgCrossfadeMs);
  }

  setTrackContent(track, background, options = {}) {
    if (!background || !background.path) {
      track.img.removeAttribute('src');
      track.video.removeAttribute('src');
      track.video.pause();
      track.video.style.display = 'none';
      track.img.style.display = 'none';
      track.type = 'none';
      return;
    }
    if (background.type === 'video') {
      track.video.loop = background.loop !== false;
      const forceReload = options.forceReload === true;
      if (forceReload || track.type !== 'video' || track.src !== background.path) {
        track.video.oncanplay = () => {
          track.video.play().catch(() => {});
        };
        track.video.src = background.path;
        track.video.load();
        track.video.play().catch(() => {});
      } else {
        track.video.play().catch(() => {});
      }
      track.video.style.display = 'block';
      track.img.style.display = 'none';
      track.type = 'video';
      track.src = background.path;
    } else {
      if (track.type !== 'image' || track.src !== background.path) {
        track.img.src = background.path;
      }
      track.img.style.display = 'block';
      track.video.style.display = 'none';
      track.video.pause();
      track.type = 'image';
      track.src = background.path;
    }
  }

  stopTrack(track) {
    if (track.type === 'video') {
      track.video.pause();
      track.video.removeAttribute('src');
    }
    track.type = null;
    track.src = null;
  }

  async transitionText(text, panic) {
    const token = ++this.textToken;
    this.textLayer.style.transition = `opacity ${this.textFadeMs}ms ease`;
    this.textLayer.style.opacity = '0';

    await new Promise((resolve) => window.setTimeout(resolve, this.textFadeMs));

    if (token !== this.textToken) {
      return;
    }

    this.applyText(text);
    this.applyAutoFit();

    this.isPanic = Boolean(panic);
    this.textLayer.style.opacity = this.isPanic ? '0' : '1';
  }

  setPanic(isPanic) {
    this.isPanic = Boolean(isPanic);
    this.textLayer.style.transition = `opacity ${this.textFadeMs}ms ease`;
    this.textLayer.style.opacity = this.isPanic ? '0' : '1';
  }

  applyText(text = {}) {
    const title = text.title || '';
    const lyrics = text.lyrics || '';
    const footer = text.footer || '';

    this.titleEl.textContent = title;
    this.lyricsEl.textContent = lyrics;
    this.footerEl.textContent = footer;

    this.titleEl.style.display = text.showTitle ? 'block' : 'none';
    this.lyricsEl.style.display = text.showLyrics ? 'block' : 'none';
    this.footerEl.style.display = text.showFooter ? 'block' : 'none';
    this.root.classList.toggle('no-lyrics', !text.showLyrics);
  }

  applyAutoFit() {
    if (!this.lyricsEl || this.lyricsEl.style.display === 'none') {
      this.root.style.setProperty('--lyrics-block-height', '0px');
      return;
    }
    const box = this.textBox;
    const lyrics = this.lyricsEl;
    let size = this.baseFontPx;
    lyrics.style.fontSize = `${size}px`;
    const min = this.minFontPx || 36;

    for (let i = 0; i < 40; i += 1) {
      if (lyrics.scrollHeight <= box.clientHeight && lyrics.scrollWidth <= box.clientWidth) {
        break;
      }
      size -= 2;
      if (size <= min) {
        size = min;
        break;
      }
      lyrics.style.fontSize = `${size}px`;
    }
    const height = lyrics.getBoundingClientRect().height;
    this.root.style.setProperty('--lyrics-block-height', `${height}px`);
  }
}
