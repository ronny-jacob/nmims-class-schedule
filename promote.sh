#!/usr/bin/env bash
# Promote the current staging branch to the production repo (live site).
# Run from a clone of the staging repo while on the branch you want to ship.
#
#   ./promote.sh            # push current branch -> production main
#   ./promote.sh <branch>   # push a specific branch   -> production main
#
set -euo pipefail

PROD_REPO="ronny-jacob/nmims-class-schedule"
PROD_URL="https://github.com/${PROD_REPO}.git"
BRANCH="${1:-$(git branch --show-current)}"

if [[ -z "$BRANCH" || "$BRANCH" == "HEAD" ]]; then
  echo "❌ Could not determine branch to promote. Pass one: ./promote.sh <branch>" >&2
  exit 1
fi

echo "▶ Promoting branch '$BRANCH' → $PROD_REPO (main)"
read -p "Confirm pushing to the LIVE site? [y/N] " -n 1 -r
echo
if [[ ! $REPLY =~ ^[Yy]$ ]]; then
  echo "Cancelled."
  exit 1
fi

git fetch "$PROD_URL" main
git push "$PROD_URL" "$BRANCH:main"
echo "✅ Promoted. The production Pages deploy will run now:"
echo "   https://github.com/${PROD_REPO}/actions"
