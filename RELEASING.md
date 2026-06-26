# Releasing Access Explorer

This project uses a two-layer release flow:

1. Local preflight (developer machine)
2. Safe CI preflight + optional Marketplace publish (GitHub Actions)

## 1) Local preflight

Run before opening a release PR or pushing a version bump:

```powershell
npm ci
npm run release:preflight
```

This executes:

- TypeScript compile
- Release smoke checks (`scripts/smoke-release-checks.js`)
- VSIX packaging to `.artifacts/access-explorer.vsix`

## 2) Version bump

Update these files together:

- `package.json` -> `version`
- `package-lock.json` -> root versions
- `CHANGELOG.md` -> new version section

## 3) Publish options

### Automatic publish on push to `master`

The workflow publishes only when:

- version changed, and
- push is safe (no workflow file changes in that push), and
- preflight checks pass.

### Manual run (`workflow_dispatch`)

You can run the workflow manually with:

- `publish=false`: only secure preflight checks (recommended before release)
- `publish=true`: preflight + Marketplace publish
- `expected_version` (optional): fail if `package.json` version does not match this value

## 4) Secrets

Required repository secret:

- `VSCE_PAT`

## 5) Post-publish checks

- `npx @vscode/vsce show luna-soft.access-explorer`
- verify Marketplace page version
- verify GitHub Release/tag alignment
