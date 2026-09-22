#!/bin/bash
echo "Deploying SVGs..."
git add -u
git add public/images/kapak/*.svg
git commit -m "feat: replace all raster images with 124 uniquely generated 4K pristine SVG dashboard components"
git push
npm run build && firebase deploy --only hosting
echo "Deploy Done!"
