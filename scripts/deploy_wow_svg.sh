#!/bin/bash
echo "Deploying WOW SVGs..."
git add public/images/kapak/*.svg
git commit -m "feat: upgrade all product covers to Enterprise Intelligence (Palantir/Primer) deep web 4K generative SVGs"
git push
npm run build && firebase deploy --only hosting
echo "Deploy Done!"
