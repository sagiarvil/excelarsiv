#!/bin/bash
echo "Waiting for images to finish processing..."
while pgrep -f "premium_upscale.mjs" > /dev/null; do
  sleep 2
done
echo "Images processed. Running build and deploy..."
npm run build && git add public/images/kapak/*.png && git commit -m "feat: replace low-res images with ultra-premium 4K color-shifted templates" && git push && firebase deploy --only hosting
echo "Done!"
