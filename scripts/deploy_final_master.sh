#!/bin/bash
set -e
echo "Starting final deploy of 124 redrawn master vector covers..."
git add public/images/kapak/*.svg src/lib/kapak.ts
git commit -m "feat: redraw all 124 product covers as unique, flawless light-theme vector infographics from urun-gorselleri examples"
git push origin main
npm run build
firebase deploy --only hosting
echo "FINAL DEPLOY SUCCESS!"
