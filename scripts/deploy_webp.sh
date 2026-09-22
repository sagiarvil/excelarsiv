#!/bin/bash
echo "Building and deploying WebP..."
npm run build && git add src/lib/kapak.ts public/images/kapak/*.webp && git commit -m "feat: use highly optimized 1280x800 WebP images directly mapped to 124 unique products for flawless quality without pixelation" && git push && firebase deploy --only hosting
echo "Done!"
