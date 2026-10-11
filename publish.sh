#!/bin/sh
# Copy the latest dashboard from the parent folder into this site and push it to GitHub Pages.
set -e
cd "$(dirname "$0")"
cp ../Energy_Management_Market_Map.html index.html
cp ../apps-script/Code.gs apps-script/Code.gs
cp ../energy_management_companies.csv energy_management_companies.csv
git add -A
git commit -m "Update market map" || { echo "Nothing to publish."; exit 0; }
git push
echo "Published — live in about a minute at https://sayasivaram.github.io/energy-management-market-map/"
