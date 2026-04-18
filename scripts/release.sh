#!/bin/bash
# ── BPHH DocGen Release Script ────────────────────────────────────────────────
# Mặc định: build Windows EXE.  Dùng --mac để build thêm macOS DMG.
#
# Usage:
#   cd bphh_docgen
#   ./scripts/release.sh                       # Build + release Windows
#   ./scripts/release.sh --win-only            # Build + release Windows only
#   ./scripts/release.sh --mac                 # Build + release Windows + Mac
#   ./scripts/release.sh --mac-only            # Build + release Mac only
#   ./scripts/release.sh --no-build            # Dùng artifact có sẵn
#   ./scripts/release.sh --bump=patch          # Tăng patch trước khi release
#   ./scripts/release.sh --bump=minor --mac    # Tăng minor + Win + Mac
#
# Prerequisites:
#   - gh CLI installed & authenticated: gh auth login
#   - SSH key configured for git@github.com:phongbksneep/toolgendocBPH.git

set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ELECTRON_DIR="$(dirname "$SCRIPT_DIR")"
ROOT_DIR="$(dirname "$ELECTRON_DIR")"

cd "$ELECTRON_DIR"

# ── Parse args ────────────────────────────────────────────────────────────────
BUILD_WIN=true    # Mặc định: build Windows
BUILD_MAC=false   # Mac là tuỳ chọn (--mac)
NO_BUILD=false
BUMP=""
for arg in "$@"; do
  case $arg in
    --mac)        BUILD_MAC=true ;;
    --mac-only)   BUILD_MAC=true; BUILD_WIN=false ;;
    --win-only)   BUILD_WIN=true; BUILD_MAC=false ;;
    --no-build)   NO_BUILD=true ;;
    --bump=*)     BUMP="${arg#*=}" ;;
  esac
done

# ── Read / bump version ───────────────────────────────────────────────────────
CURRENT_VERSION=$(node -p "require('./package.json').version")
echo "Current version: $CURRENT_VERSION"

if [ -n "$BUMP" ]; then
  npx --yes semver -i "$BUMP" "$CURRENT_VERSION" > /tmp/newver.txt
  NEW_VERSION=$(cat /tmp/newver.txt)
  # Bump package.json
  node -e "
    const pkg = require('./package.json');
    pkg.version = '$NEW_VERSION';
    require('fs').writeFileSync('./package.json', JSON.stringify(pkg, null, 2) + '\n');
  "
  # Bump main.cjs APP_VERSION (GNU/BSD sed compatible)
  sed -i.bak "s/const APP_VERSION = '[^']*'/const APP_VERSION = '$NEW_VERSION'/" electron/main.cjs
  rm -f electron/main.cjs.bak
  echo "Bumped to: $NEW_VERSION"
else
  NEW_VERSION="$CURRENT_VERSION"
fi

TAG="v$NEW_VERSION"
DATE=$(date +%Y-%m-%d)

# ── Build ─────────────────────────────────────────────────────────────────────
if [ "$NO_BUILD" = true ]; then
  echo "⏭ Skipping build (--no-build), using existing artifacts in release/"
else
  if [ "$BUILD_MAC" = true ]; then
    echo ""
    echo "▶ Building macOS..."
    npm run build:mac
  fi

  if [ "$BUILD_WIN" = true ]; then
    echo ""
    echo "▶ Building Windows..."
    npm run build:win
  fi
fi

RELEASE_DIR="$ELECTRON_DIR/release"
DMG_ARM64="$RELEASE_DIR/BPHH-DocGen-${NEW_VERSION}-arm64.dmg"
DMG_X64="$RELEASE_DIR/BPHH-DocGen-${NEW_VERSION}.dmg"
# Windows: installer + portable zip (no spaces in filenames via artifactName config)
WIN_EXE="$RELEASE_DIR/BPHH-DocGen-Setup-${NEW_VERSION}.exe"
WIN_ZIP="$RELEASE_DIR/BPHH-DocGen-${NEW_VERSION}-x64.zip"

# ── Update version.json in root repo ─────────────────────────────────────────
echo ""
echo "▶ Updating version.json..."
VERSION_JSON="$ROOT_DIR/version.json"

node -e "
const v = '$NEW_VERSION', tag = '$TAG', date = '$DATE';
const base = 'https://github.com/phongbksneep/toolgendocBPH/releases/download/' + tag;
// Dùng tên file không có dấu cách (artifactName đã cấu hình trong package.json)
const obj = {
  version: v,
  date,
  changelog: process.env.CHANGELOG || 'Xem chi tiết: https://github.com/phongbksneep/toolgendocBPH/releases/tag/' + tag,
  mac_arm64:  base + '/BPHH-DocGen-' + v + '-arm64.dmg',
  mac_x64:    base + '/BPHH-DocGen-' + v + '.dmg',
  win_x64:    base + '/BPHH-DocGen-Setup-' + v + '.exe',
  win_x64_zip: base + '/BPHH-DocGen-' + v + '-x64.zip',
};
require('fs').writeFileSync('$VERSION_JSON', JSON.stringify(obj, null, 2) + '\n');
console.log('Written: $VERSION_JSON');
"

# ── Push version.json to toolgendocBPH GitHub repo ───────────────────────────
echo ""
echo "▶ Pushing version.json to GitHub..."

REMOTE_REPO="git@github.com:phongbksneep/toolgendocBPH.git"
TMPDIR_CLONE=$(mktemp -d)

git clone --depth 1 "$REMOTE_REPO" "$TMPDIR_CLONE" 2>/dev/null || {
  echo "⚠ Clone failed. Initializing fresh repo..."
  mkdir -p "$TMPDIR_CLONE"
  cd "$TMPDIR_CLONE"
  git init
  git remote add origin "$REMOTE_REPO"
}

cp "$VERSION_JSON" "$TMPDIR_CLONE/version.json"

cd "$TMPDIR_CLONE"
git add version.json
git diff --cached --quiet && echo "version.json unchanged, skipping push." || {
  git config user.email "deploy@bphh" 2>/dev/null || true
  git config user.name  "BPHH Deploy" 2>/dev/null || true
  git commit -m "release: $TAG ($DATE)"
  git push origin HEAD:main
  echo "✅ version.json pushed to GitHub"
}

cd "$ELECTRON_DIR"
rm -rf "$TMPDIR_CLONE"

# ── Create GitHub Release and upload assets ───────────────────────────────────
if command -v gh &>/dev/null; then
  echo ""
  echo "▶ Creating GitHub Release $TAG..."

  # Change to root dir which has the same git remote
  cd "$ROOT_DIR"

  ASSETS=()
  [ -f "$DMG_ARM64" ] && ASSETS+=("$DMG_ARM64")
  [ -f "$DMG_X64"   ] && ASSETS+=("$DMG_X64")
  # Windows: ưu tiên EXE installer, kèm ZIP portable
  [ -f "$WIN_EXE"   ] && ASSETS+=("$WIN_EXE")
  [ -f "$WIN_ZIP"   ] && ASSETS+=("$WIN_ZIP")

  # Demo / sample files
  DEMO_JSON="$ROOT_DIR/project-data.demo.friendly.json"
  DEMO_XLSX="$ROOT_DIR/assets/project-data-mau.xlsx"
  MANUAL_MD="$ROOT_DIR/HUONG-DAN-SU-DUNG.md"
  MANUAL_DOCX="$ROOT_DIR/HUONG-DAN-SU-DUNG.docx"
  # Convert manual MD → DOCX if pandoc available
  if command -v pandoc &>/dev/null && [ -f "$MANUAL_MD" ]; then
    pandoc "$MANUAL_MD" -o "$MANUAL_DOCX" --from markdown --to docx -V geometry:margin=2cm 2>/dev/null && echo "✅ Manual converted to DOCX"
  fi
  [ -f "$DEMO_JSON"    ] && ASSETS+=("$DEMO_JSON")
  [ -f "$DEMO_XLSX"    ] && ASSETS+=("$DEMO_XLSX")
  [ -f "$MANUAL_DOCX" ] && ASSETS+=("$MANUAL_DOCX")

  NOTES="${CHANGELOG:-Phiên bản $NEW_VERSION}"

  # Delete existing release if present (idempotent re-deploy)
  gh release delete "$TAG" --repo "phongbksneep/toolgendocBPH" --yes 2>/dev/null || true

  gh release create "$TAG" \
    --repo "phongbksneep/toolgendocBPH" \
    --title "BPHH-DocGen $NEW_VERSION" \
    --notes "$NOTES" \
    --draft=false \
    "${ASSETS[@]}" && echo "✅ GitHub Release created: $TAG" \
    || { echo "❌ gh release create failed"; exit 1; }
else
  echo ""
  echo "⚠ gh CLI not found. Install: brew install gh"
  echo "   Then manually:"
  echo "   1. Go to https://github.com/phongbksneep/toolgendocBPH/releases/new"
  echo "   2. Tag: $TAG"
  echo "   3. Upload: $DMG_ARM64"
  echo "              $DMG_X64"
  [ "$BUILD_WIN" = true ] && echo "              $WIN_EXE"
fi

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "✅ Release $TAG complete"
echo "   version.json → https://raw.githubusercontent.com/phongbksneep/toolgendocBPH/main/version.json"
echo "   Releases     → https://github.com/phongbksneep/toolgendocBPH/releases"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
