#!/usr/bin/env bash
#
# Prepare the files in this repository for releasing a new version.
#
#   prepare-release.sh 2.0.0
#   prepare-release.sh --notes /tmp/notes.md 2.0.0
#
# This edits the working tree and does nothing else: it doesn't commit, tag,
# push, or talk to GitHub. .github/workflows/release.yml runs it and does the
# rest. Running it by hand shows what a release would change; "git diff"
# shows the result and "git checkout ." undoes it.
#
# It checks that VERSION has the form X.Y.Z, that it is greater than the
# version in codemeta.json, that tag vVERSION doesn't exist, and that the
# [Unreleased] section of CHANGELOG.md has something in it. Then it
#
#  - sets "version" and "datePublished" in codemeta.json;
#  - copies the version, the date, and the title, abstract, URLs, and license
#    from codemeta.json into CITATION.cff;
#  - turns the [Unreleased] section of CHANGELOG.md into a section for
#    VERSION, starts a new empty [Unreleased] section, and updates the compare
#    links at the bottom;
#  - with --notes, writes the new CHANGELOG.md section to a file, for use as
#    the release notes.
#
# codemeta.json is edited line by line rather than rewritten by jq, so that
# the diff touches only the values that changed.
#
# Needs bash, git, jq, and awk.
#
# Copyright 2026 California Institute of Technology.
# License: Modified BSD 3-clause – see file "LICENSE" in the project website.
# Website: https://github.com/caltechlibrary/boffo

set -euo pipefail

VERSION=""
NOTES=""
TODAY="$(date +%F)"

usage() {
  cat <<'USAGE'
prepare-release.sh -- update codemeta.json, CITATION.cff, and CHANGELOG.md
for a new release

Usage: prepare-release.sh [options] VERSION

Options:
  --notes FILE    write the new CHANGELOG.md section to FILE
  --date DATE     release date, YYYY-MM-DD           (default: today)
  -h, --help      this text
USAGE
}

die() { echo "prepare-release: $*" >&2; exit 1; }

while [ $# -gt 0 ]; do
  case "$1" in
    --notes)    NOTES="$2"; shift 2 ;;
    --date)     TODAY="$2"; shift 2 ;;
    -h|--help)  usage; exit 0 ;;
    -*)         echo "prepare-release: unknown option: $1" >&2; usage >&2; exit 2 ;;
    *)          [ -z "$VERSION" ] || { usage >&2; exit 2; }; VERSION="$1"; shift ;;
  esac
done

[ -n "$VERSION" ] || { usage >&2; exit 2; }
for p in git jq awk; do
  command -v "$p" >/dev/null || die "$p is not installed"
done

cd "$(git rev-parse --show-toplevel)"
for f in codemeta.json CITATION.cff CHANGELOG.md; do
  [ -f "$f" ] || die "no $f in $(pwd)"
done

# Check the request ──────────────────────────────────────────────────────────

[[ "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]] \
  || die "version must look like 1.2.3, not \"$VERSION\""
[[ "$TODAY" =~ ^[0-9]{4}-[0-9]{2}-[0-9]{2}$ ]] \
  || die "date must look like 2026-01-31, not \"$TODAY\""

CURRENT="$(jq -r .version codemeta.json)"
newer="$(printf '%s\n%s\n' "$CURRENT" "$VERSION" | sort -V | tail -1)"
if [ "$VERSION" = "$CURRENT" ] || [ "$newer" != "$VERSION" ]; then
  die "version $VERSION is not greater than $CURRENT in codemeta.json"
fi

TAG="v$VERSION"
if git rev-parse -q --verify "refs/tags/$TAG" >/dev/null; then
  die "tag $TAG already exists"
fi

# Lines of the [Unreleased] section, up to the next version heading.
unreleased="$(awk '
  /^## \[Unreleased\]/ { inside = 1; next }
  inside && /^## \[/   { exit }
  inside               { print }
' CHANGELOG.md)"
grep -q '^## \[Unreleased\]' CHANGELOG.md \
  || die "CHANGELOG.md has no \"## [Unreleased]\" heading"
[ -n "$(tr -d '[:space:]' <<< "$unreleased")" ] \
  || die "the [Unreleased] section of CHANGELOG.md is empty"

# The [Unreleased] compare link names the previous tag and the repository.
link="$(grep -E '^\[Unreleased\]: ' CHANGELOG.md || true)"
[[ "$link" =~ ^\[Unreleased\]:\ (.*)/compare/(.*)\.\.\.HEAD$ ]] \
  || die "CHANGELOG.md needs a link like \"[Unreleased]: <repo>/compare/v1.2.3...HEAD\""
REPO_URL="${BASH_REMATCH[1]}"
PREV_TAG="${BASH_REMATCH[2]}"

# Edit the files ─────────────────────────────────────────────────────────────

# Replace one file with the output of a command run on it.
rewrite() {
  local file="$1"; shift
  local tmp
  tmp="$(mktemp)"
  "$@" "$file" > "$tmp"
  cat "$tmp" > "$file"
  rm -f "$tmp"
}

# codemeta.json: only the top-level keys, which are indented by two spaces.
# programmingLanguage has a "version" of its own one level further in.
rewrite codemeta.json awk -v version="$VERSION" -v date="$TODAY" '
  /^  "version": /       { sub(/: ".*"/, ": \"" version "\"") }
  /^  "datePublished": / { sub(/: ".*"/, ": \"" date "\"") }
  { print }
'
[ "$(jq -r .version codemeta.json)" = "$VERSION" ] \
  || die "could not set version in codemeta.json"
[ "$(jq -r .datePublished codemeta.json)" = "$TODAY" ] \
  || die "could not set datePublished in codemeta.json"

# CITATION.cff: codemeta.json is the source of these values. They go through
# ENVIRON so that characters like "&" and "\" reach the file unchanged.
cff_set() {
  local key="$1" value="$2"
  [ -n "$value" ] && [ "$value" != null ] || return 0
  grep -q "^$key:" CITATION.cff || die "CITATION.cff has no \"$key:\" line"
  # shellcheck disable=SC2016 # $0 is awk's, not the shell's.
  VALUE="$value" rewrite CITATION.cff awk -v key="$key" '
    index($0, key ":") == 1 { print key ": " ENVIRON["VALUE"]; next }
    { print }
  '
}
cff_set title           "$(jq -r .name codemeta.json)"
cff_set abstract        "$(jq -r .description codemeta.json)"
cff_set version         "$VERSION"
cff_set date-released   "$TODAY"
cff_set url             "$(jq -r '.url // .codeRepository' codemeta.json)"
cff_set repository-code "$(jq -r .codeRepository codemeta.json)"
cff_set license-url     "$(jq -r .license codemeta.json)"

# CHANGELOG.md: the old [Unreleased] heading becomes the version's heading,
# under a new empty [Unreleased], and the version gets a compare link.
rewrite CHANGELOG.md awk \
  -v version="$VERSION" -v date="$TODAY" \
  -v repo="$REPO_URL" -v prev="$PREV_TAG" -v tag="$TAG" '
  /^## \[Unreleased\]/ {
    print "## [Unreleased]"
    print ""
    print ""
    print "## [" version "] - " date
    next
  }
  /^\[Unreleased\]: / {
    print "[Unreleased]: " repo "/compare/" tag "...HEAD"
    print "[" version "]: " repo "/compare/" prev "..." tag
    next
  }
  { print }
'

if [ -n "$NOTES" ]; then
  # The section's body, without its heading or surrounding blank lines.
  awk -v heading="## [$VERSION] - $TODAY" '
    $0 == heading      { inside = 1; next }
    inside && /^## \[/ { exit }
    inside             { lines[++n] = $0 }
    END {
      first = 1; while (first <= n && lines[first] ~ /^[[:space:]]*$/) first++
      last = n;  while (last >= first && lines[last] ~ /^[[:space:]]*$/) last--
      for (i = first; i <= last; i++) print lines[i]
    }
  ' CHANGELOG.md > "$NOTES"
  [ -s "$NOTES" ] || die "could not extract release notes from CHANGELOG.md"
fi

echo "prepare-release: $CURRENT → $VERSION ($TODAY), previous tag $PREV_TAG"
