#!/bin/bash
echo "Deploying Modern Light SVGs..."
git rm -r --cached public/images/kapak/*.webp || true
rm -f public/images/kapak/*.webp
git add public/images/kapak/*.svg
git add src/lib/kapak.ts
git commit -m "feat: replace covers with flawless 2560x1600 Light Theme Enterprise Excel vector graphics"
git push
npm run build && firebase deploy --only hosting
echo "Deploy Done!"
