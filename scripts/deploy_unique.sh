#!/bin/bash
echo "Waiting for images to finish processing..."
while pgrep -f "premium_upscale_unique.mjs" > /dev/null; do
  sleep 2
done
echo "Images processed. Running build and deploy..."
npm run build && git add public/images/kapak/*.png && git commit -m "fix: enforce 100% unique 4K product images per user feedback" && git push && firebase deploy --only hosting
echo "Done!"
