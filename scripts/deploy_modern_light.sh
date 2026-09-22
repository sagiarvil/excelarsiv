#!/bin/bash
echo "Deploying Modern Light Showcase..."
git rm -r --cached public/images/kapak/*.svg || true
rm -f public/images/kapak/*.svg
git add public/images/kapak/*.webp
git add src/lib/kapak.ts
git commit -m "feat: redesign product covers using original screenshots wrapped in pristine 2560x1600 Apple-style light theme glass cards"
git push
npm run build && firebase deploy --only hosting
echo "Deploy Done!"
