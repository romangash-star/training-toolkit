#!/usr/bin/env bash
# Publish the site to both addresses:
#   origin -> https://www.romanai.net            (custom domain, repo training-toolkit)
#   mirror -> https://romangash-star.github.io/toolkit/   (backup address, same files without CNAME)
# The backup exists because some office firewalls block newly registered domains.
set -euo pipefail
cd "$(dirname "$0")/.."
git push -q origin main
tmp=$(mktemp); trap 'rm -f "$tmp"' EXIT
GIT_INDEX_FILE="$tmp" git read-tree HEAD
GIT_INDEX_FILE="$tmp" git rm -q --cached CNAME
tree=$(GIT_INDEX_FILE="$tmp" git write-tree)
commit=$(git -c user.name=ROMAN -c user.email=roman.gash@gmail.com commit-tree "$tree" -m "Mirror of $(git rev-parse --short HEAD) without the custom domain")
git push -q -f mirror "$commit:refs/heads/main"
echo "published: romanai.net and the github.io backup"
