#!/usr/bin/env bash
# Kullanım: ./set-host.sh https://yeni-adres.example.com
# manifest.xml ve config.js içindeki HOST adresini değiştirir (sondaki / atılır).
set -euo pipefail
NEW="${1:?https://... adres verin}"; NEW="${NEW%/}"
cd "$(dirname "$0")"
OLD=$(sed -n 's/.*HOST_URL: "\([^"]*\)".*/\1/p' config.js)
[ -n "$OLD" ] || { echo "config.js içinde HOST_URL bulunamadı"; exit 1; }
sed -i '' "s|$OLD|$NEW|g" manifest.xml config.js
echo "$OLD -> $NEW"; grep -c "$NEW" manifest.xml config.js
