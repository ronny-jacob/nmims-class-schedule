#!/usr/bin/env bash
# ⚠️  DEPRECATED — Use the GitHub Actions workflow instead.
#
# Use the workflow on the staging repo:
#   Repository: ronny-jacob/nmims-class-schedule-staging
#   Workflow:   "Promote to Production" (Actions → run workflow)
#
# Why the workflow is preferred:
#   - It refuses to promote when production is ahead of staging.
#   - It refuses when staging and production are identical (no-op guard).
#   - It runs through GitHub's audit log + permissions model.
#   - This script does not have those checks and silently force-pushes.
#
# Delete this file once the workflow is the established source of truth.

set -euo pipefail

PROD_REPO="ronny-jacob/nmims-class-schedule"
PROD_URL="https://github.com/${PROD_REPO}.git"
BRANCH="${1:-$(git branch --show-current)}"

echo "❌ promote.sh is deprecated. Use the 'Promote to Production' workflow."
echo "   https://github.com/${PROD_REPO}/actions"
exit 1
