# Worship Presenter

Offline Windows worship lyrics presenter.

## Install (Windows)

1. Open the GitHub repo page and go to Releases.
2. Open the latest published release (not a draft).
3. Download `worship-presenter-<version> Setup.exe` from Assets.
4. Run the installer (Windows SmartScreen: More info -> Run anyway).
5. Launch "Worship Presenter" from the Start Menu.

## Updates

- In the app: File -> Check for Updates.
- Auto-updates only work from published releases in a public GitHub repo.

## Build and Publish (maintainers)

1. Ensure Node.js and npm are installed.
2. Set `GITHUB_TOKEN` with repo access in your environment.
3. Bump `version` in `package.json` (release tags should be `vX.Y.Z`).
4. Run `npm run publish`.
5. Find artifacts in `out-publish\make`.
6. In GitHub Releases, publish the draft release so auto-updates work.
