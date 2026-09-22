#!/bin/bash
while pgrep -f "deploy_modern_light.sh" > /dev/null; do
  sleep 2
done
echo "Catchup commit for remaining images..."
git add public/images/kapak/*.webp
if ! git diff --cached --quiet; then
  git commit -m "chore: commit remaining generated showcase webp images"
  git push
  npm run build && firebase deploy --only hosting
fi
echo "All done!"
