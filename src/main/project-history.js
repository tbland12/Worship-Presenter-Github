const crypto = require('node:crypto');
const path = require('node:path');

const MAX_RECENT_PROJECTS = 50;

function projectHistoryId(filePath) {
  return crypto.createHash('sha256').update(path.resolve(filePath).toLowerCase()).digest('hex').slice(0, 32);
}

function countSlides(project) {
  const songSlides = Object.values(project?.songs || {}).reduce((total, song) => {
    return total + (Array.isArray(song?.slides) ? song.slides.length : 0);
  }, 0);
  const announcementSlides = Array.isArray(project?.announcements?.slides)
    ? project.announcements.slides.length
    : 0;
  const timerSlides = Array.isArray(project?.timer?.slides) ? project.timer.slides.length : 0;
  return songSlides + announcementSlides + timerSlides;
}

function buildHistoryEntry(filePath, project, savedAt = new Date().toISOString()) {
  const songs = Object.values(project?.songs || {}).map((song) => ({
    id: String(song?.id || ''),
    title: String(song?.title || 'Untitled Song'),
    ccliSongNumber: String(song?.ccli?.songNumber || '')
  }));
  return {
    id: projectHistoryId(filePath),
    filePath: path.resolve(filePath),
    title: path.basename(filePath, path.extname(filePath)) || 'Untitled Session',
    lastSavedAt: savedAt,
    slideCount: countSlides(project),
    songs
  };
}

function upsertHistory(history, entry) {
  const entries = [entry, ...(history?.entries || []).filter((item) => item.id !== entry.id)]
    .slice(0, MAX_RECENT_PROJECTS);
  return { version: 1, entries };
}

function removeHistoryEntry(history, id) {
  return {
    version: 1,
    entries: (history?.entries || []).filter((entry) => entry.id !== id)
  };
}

function publicHistoryEntries(history, limit = 10) {
  return (history?.entries || []).slice(0, limit).map(({ filePath: _filePath, ...entry }) => entry);
}

module.exports = {
  MAX_RECENT_PROJECTS,
  buildHistoryEntry,
  countSlides,
  projectHistoryId,
  publicHistoryEntries,
  removeHistoryEntry,
  upsertHistory
};
