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

1. Install Node.js 22.12 or newer.
2. Run `npm ci` and `npm run check`.
3. Set `GITHUB_TOKEN` with repo access in your environment.
4. Bump `version` in `package.json` (release tags should be `vX.Y.Z`).
5. Run `npm run publish`.
6. Find artifacts in `out-publish\make`.
7. In GitHub Releases, publish the draft release so auto-updates work.

## Release Checklist

- [ ] Update `package.json` version and commit the change.
- [ ] Run `npm run publish`.
- [ ] Confirm assets are attached to the `vX.Y.Z` release.
- [ ] Publish the release (not draft).
- [ ] Install the new Setup.exe and verify `File -> Check for Updates`.
