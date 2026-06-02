#!/bin/bash
# Example deploy: push public/ to a server web root and set the add-in host.
# Usage: ./deploy.sh archive.example.com root@server /var/www/outlook-openarchiver/public
set -e
HOST="${1:?domain, e.g. archive.example.com}"
TARGET="${2:?ssh target, e.g. root@server}"
WEBROOT="${3:?web root path on server}"
tmp=$(mktemp -d); cp -r public/* "$tmp/"
sed -i "s/ARCHIV_HOST/$HOST/g" "$tmp/manifest.xml"
ssh "$TARGET" "mkdir -p $WEBROOT"
scp -r "$tmp"/* "$TARGET:$WEBROOT/"
rm -rf "$tmp"
echo "Deployed. Manifest URL: https://$HOST/manifest.xml"
