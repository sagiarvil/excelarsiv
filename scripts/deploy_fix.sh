#!/bin/bash
while pgrep -f "deploy_webp.sh" > /dev/null; do
  sleep 2
done
echo "Running final build and deploy to sync provenance..."
npm run build && firebase deploy --only hosting
echo "Done!"
