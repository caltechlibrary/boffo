#!/usr/bin/env bash
#
# Record the DOI of the newest archived release in CITATION.cff and
# codemeta.json.
#
#   update-doi.sh
#   update-doi.sh --tries 10 --wait 30
#
# The "relatedLink" field of codemeta.json holds the URL of an archived
# release in an InvenioRDM server such as CaltechDATA. InvenioRDM can report
# the latest version of any record, so this asks it for the latest version
# of that one, and then sets "doi" in CITATION.cff and "relatedLink" in
# codemeta.json to point to it.
#
# The latest record has to be for the version in codemeta.json. Right after
# archiving, InvenioRDM may still report the previous release for a little
# while, so this asks again (--tries times, --wait seconds apart) before
# giving up. Writing the previous release's DOI would make CITATION.cff cite
# the wrong version.
#
# Run it after archiving a release with the InvenioRDM GitHub Archiver
# workflow (iga.yml), then commit the result. It edits the working tree and
# does nothing else. See "Archiving a release" in docs/releasing.md.
#
# Needs bash, git, curl, jq, and awk.
#
# Copyright 2026 California Institute of Technology.
# License: Modified BSD 3-clause – see file "LICENSE" in the project website.
# Website: https://github.com/caltechlibrary/boffo

set -euo pipefail

TRIES=10
WAIT=30

usage() {
  cat <<'USAGE'
update-doi.sh -- record the newest archived release's DOI in CITATION.cff
and codemeta.json

Options:
  --tries N    times to ask InvenioRDM for the record      (default: 10)
  --wait S     seconds to wait between tries               (default: 30)
  -h, --help   this text
USAGE
}

die() { echo "update-doi: $*" >&2; exit 1; }

while [ $# -gt 0 ]; do
  case "$1" in
    --tries)    TRIES="$2"; shift 2 ;;
    --wait)     WAIT="$2"; shift 2 ;;
    -h|--help)  usage; exit 0 ;;
    *)          echo "update-doi: unknown option: $1" >&2; usage >&2; exit 2 ;;
  esac
done

for p in git curl jq awk; do
  command -v "$p" >/dev/null || die "$p is not installed"
done

cd "$(git rev-parse --show-toplevel)"
for f in codemeta.json CITATION.cff; do
  [ -f "$f" ] || die "no $f in $(pwd)"
done

VERSION="$(jq -r .version codemeta.json)"
LINK="$(jq -r '.relatedLink | strings' codemeta.json)"
[[ "$LINK" =~ ^(https?://[^/]+)/records/([^/]+)$ ]] \
  || die "relatedLink in codemeta.json must be one URL like https://server/records/<id>, not \"$LINK\""
SERVER="${BASH_REMATCH[1]}"
RECORD="${BASH_REMATCH[2]}"
API="$SERVER/api/records/$RECORD/versions/latest"

for ((try = 1; ; try++)); do
  latest="$(curl -fsSL "$API")" || die "could not get $API"
  latest_version="$(jq -r '.metadata.version // empty' <<< "$latest")"
  [ "$latest_version" = "$VERSION" ] && break
  if [ "$try" -ge "$TRIES" ]; then
    die "the latest record in $SERVER is for version \"$latest_version\", not $VERSION"
  fi
  echo "update-doi: latest record is for \"$latest_version\", not $VERSION; asking again in ${WAIT}s"
  sleep "$WAIT"
done

DOI="$(jq -r '.pids.doi.identifier // empty' <<< "$latest")"
ID="$(jq -r '.id // empty' <<< "$latest")"
[ -n "$DOI" ] && [ -n "$ID" ] || die "the latest record in $SERVER has no DOI or id"
NEW_LINK="$SERVER/records/$ID"

# Replace one file with the output of a command run on it.
rewrite() {
  local file="$1"; shift
  local tmp
  tmp="$(mktemp)"
  "$@" "$file" > "$tmp"
  cat "$tmp" > "$file"
  rm -f "$tmp"
}

grep -q '^doi:' CITATION.cff || die "CITATION.cff has no \"doi:\" line"
rewrite CITATION.cff awk -v doi="$DOI" '
  /^doi:/ { print "doi: " doi; next }
  { print }
'

rewrite codemeta.json awk -v link="$NEW_LINK" '
  /^  "relatedLink": / { sub(/: ".*"/, ": \"" link "\"") }
  { print }
'
[ "$(jq -r .relatedLink codemeta.json)" = "$NEW_LINK" ] \
  || die "could not set relatedLink in codemeta.json"

echo "update-doi: version $VERSION is $DOI ($NEW_LINK)"
