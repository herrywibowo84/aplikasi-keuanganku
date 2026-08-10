#!/usr/bin/env bash
# ============================================================
# deploy.sh — KeuanganKu Auto Deploy
# Usage: ./deploy.sh "Deskripsi perubahan"
# Contoh: ./deploy.sh "Tambah fitur export CSV"
# ============================================================

set -e

DEPLOYMENT_ID="AKfycbzwEYqyjw04RtD1MquvdFVv6IgQSyiKyzO7y_m_SPTeGvzztwf9GtCkvD1ghMl6CeqN"
CODE_FILE="Code.js"
HTML_FILE="Index.html"

# ── 1. Baca versi saat ini dari Code.js ─────────────────────
CURRENT=$(grep "const APP_VERSION" "$CODE_FILE" | sed "s/.*'\(.*\)'.*/\1/")
if [ -z "$CURRENT" ]; then
  echo "❌ Gagal membaca APP_VERSION dari $CODE_FILE"
  exit 1
fi

MAJOR=$(echo "$CURRENT" | cut -d'.' -f1)
MINOR=$(echo "$CURRENT" | cut -d'.' -f2)
NEW_MINOR=$((MINOR + 1))
NEW_VERSION="${MAJOR}.${NEW_MINOR}"

echo "📦 Versi saat ini : v${CURRENT}"
echo "🚀 Versi baru     : v${NEW_VERSION}"
echo ""

# ── 2. Update APP_VERSION di Code.js ────────────────────────
sed -i '' "s/const APP_VERSION = '[^']*'/const APP_VERSION = '${NEW_VERSION}'/" "$CODE_FILE"

# ── 3. Update fallback label versi di Index.html ─────────────
sed -i '' "s/v[0-9]\+\.[0-9]\+ &bull;/v${NEW_VERSION} \&bull;/" "$HTML_FILE"
sed -i '' "s/v[0-9]\+\.[0-9]\+ •/v${NEW_VERSION} •/" "$HTML_FILE"

echo "✅ Versi diperbarui di Code.js dan Index.html"

# ── 4. Push ke Google Apps Script ────────────────────────────
echo ""
echo "⬆️  Pushing ke GAS..."
clasp push --force

# ── 5. Deploy ke production ──────────────────────────────────
DESC="${1:-v${NEW_VERSION} - update}"
echo ""
echo "🌐 Deploying: ${DESC}"
clasp deploy -i "$DEPLOYMENT_ID" -d "v${NEW_VERSION} - ${DESC}"

# ── 6. Git commit ─────────────────────────────────────────────
echo ""
echo "💾 Git commit..."
git add -A
git commit -m "deploy: v${NEW_VERSION} - ${DESC}

Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>"

echo ""
echo "✅ Deploy selesai! Aplikasi sekarang di versi v${NEW_VERSION}"
echo "🔗 https://script.google.com/macros/s/${DEPLOYMENT_ID}/exec"
