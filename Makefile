# Makefile for developing Boffo.
# Run "make" or "make help" to get a list of commands in this makefile.
#
# Releases are made with GitHub Actions, not with this makefile. See the
# "Releasing Boffo" page in docs/ (docs/releasing.md).
#
# The codemeta.json file is the master source for the version number and
# other metadata. The release workflow copies them into CITATION.cff.
#
# Copyright 2024 California Institute of Technology.
# License: Modified BSD 3-clause – see file "LICENSE" in the project website.
# Website: https://github.com/caltechlibrary/boffo

SHELL=/bin/bash
.ONESHELL:                              # Run all commands in the same shell.
.SHELLFLAGS += -e                       # Exit at the first error.

# This Makefile uses syntax that needs at least GNU Make version 3.82.
# The following test is based on the approach posted by Eldar Abusalimov to
# Stack Overflow in 2012 at https://stackoverflow.com/a/12231321/743730

ifeq ($(filter undefine,$(value .FEATURES)),)
$(error Unsupported version of Make. \
    This Makefile does not work properly with GNU Make $(MAKE_VERSION); \
    it needs GNU Make version 3.82 or later)
endif

# Before we go any further, test if certain programs are available.
# The following is based on the approach posted by Jonathan Ben-Avraham to
# Stack Overflow in 2014 at https://stackoverflow.com/a/25668869
#
# Every target needs git and jq, to set the variables below. A target that
# needs another program checks for it with "require" as its first line, so
# that "make help" works without clasp, for example.

programs_needed = git jq
TEST := $(foreach p,$(programs_needed),\
	  $(if $(shell which $(p)),_,$(error Cannot find program "$(p)")))

require = $(foreach p,$(1),\
	  $(if $(shell which $(p)),,$(error Cannot find program "$(p)")))

# Set some basic variables. These are quick to set; we set additional ones
# using the dependency named "vars" but only when the others are needed.

name	 := $(strip $(shell jq -r .name codemeta.json))
progname := $(strip $(shell jq -r '.identifier | ascii_downcase' codemeta.json))
version	 := $(strip $(shell jq -r .version codemeta.json))
repo	 := $(shell git ls-remote --get-url | sed -e 's/.*:\(.*\).git/\1/')
repo_url := https://github.com/$(repo)
branch	 := $(shell git rev-parse --abbrev-ref HEAD)


# Print help if no command is given ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# The help scheme works by looking for lines beginning with "#:" above make
# targets in this file. Originally based on code posted to Stack Overflow on
# 2019-11-28 by Richard Kiefer at https://stackoverflow.com/a/59087509/743730

#: Print a summary of available commands.
help:
	@echo "This is the Makefile for $(bright)$(name)$(reset)."
	@echo "Available commands:"
	@echo
	@grep -B1 -E "^[a-zA-Z0-9_-]+\:([^\=]|$$)" $(MAKEFILE_LIST) \
	| grep -v -- -- \
	| sed 'N;s/\n/###/' \
	| sed -n 's/^#: \(.*\)###\(.*\):.*/$(color)\2$(reset):###\1/p' \
	| column -t -s '###'

# Gather additional values we sometimes need ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# These variables take longer to compute, and for some actions like "make help"
# they are unnecessary and annoying to wait for.
vars: doi-vars
	$(eval url	:= $(strip $(shell jq -r '.url // empty' codemeta.json)))
	$(eval url	:= $(or $(url),$(repo_url)))
	$(eval license	:= $(strip $(shell jq -r .license codemeta.json)))
	$(eval desc	:= $(strip $(shell jq -r .description codemeta.json)))
	$(eval author	:= \
	  $(strip $(shell jq -r '.author[0].givenName + " " + .author[0].familyName' codemeta.json)))
	$(eval email	:= $(strip $(shell jq -r .author[0].email codemeta.json)))

# If this software isn't getting archived in InvenioRDM, the next rule will
# leave rdm_id & new_doi undefined. Other rules in this makefile test for that.
.SILENT: doi-vars
doi-vars:
	$(call require,curl)
	$(eval rdm_link	:= \
	  $(strip $(shell jq -r '.relatedLink | if type == "array" then .[0] else . end' codemeta.json)))
ifneq ($(rdm_link),null)
	$(eval rdm_url	  := $(shell cut -d'/' -f 1-3 <<< $(rdm_link)))
	$(eval rdm_id	  := $(shell sed -r 's|.*/(.*)$$|\1|' <<< $(rdm_link)))
	$(eval vers_url   := $(rdm_url)/api/records/$(rdm_id)/versions/latest)
	$(eval latest_doi := $(shell curl -L -s $(vers_url) | jq -r .pids.doi.identifier))
endif

#: Print variables set in this Makefile from various sources.
.SILENT: report
report: vars
	echo "$(color)name$(reset)	 = $(name)"	  | expand -t 21
	echo "$(color)progname$(reset)	 = $(progname)"   | expand -t 21
	echo "$(color)url$(reset)	 = $(url)"	  | expand -t 21
	echo "$(color)desc$(reset)	 = $(desc)"	  | expand -t 21
	echo "$(color)version$(reset)	 = $(version)"	  | expand -t 21
	echo "$(color)author$(reset)	 = $(author)"	  | expand -t 21
	echo "$(color)email$(reset)	 = $(email)"	  | expand -t 21
	echo "$(color)license$(reset)	 = $(license)"	  | expand -t 21
	echo "$(color)url$(reset)	 = $(url)"	  | expand -t 21
	echo "$(color)repo url$(reset)	 = $(repo_url)"   | expand -t 21
	echo "$(color)branch$(reset)	 = $(branch)"	  | expand -t 21
	echo "$(color)rdm_id$(reset)	 = $(rdm_id)"	  | expand -t 21
	echo "$(color)latest_doi$(reset) = $(latest_doi)" | expand -t 21


# make watch ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#: Continually watch files & upload changes to Google GAS project.
watch:
	$(call require,clasp)
	clasp -P $(CURDIR)/.clasp.json.caltech push --watch


# make lint & make test ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# jshint and markdownlint-cli2 run through npx at pinned versions, so they
# need only Node.js, which clasp needs too. actionlint checks the GitHub
# Actions workflows; it is a separate program, so it runs only if installed.

#: Check Code.js, the Markdown files, and (if installed) the workflows.
lint:
	@$(call require,npx)
	npx --yes jshint@2.13.6 Code.js
	npx --yes markdownlint-cli2@0.24.0 $(shell git ls-files '*.md')
	if command -v actionlint > /dev/null; then
	  actionlint
	else
	  echo "actionlint isn't installed, so the workflows weren't checked."
	fi


# Cleanup and miscellaneous directives ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#: Clean this directory of temporary and backup files.
clean:
	rm -f *.bak
	@echo ✨ Cleaned! ✨

#: Print a random joke from https://icanhazdadjoke.com/.
joke:
	$(call require,curl)
	@echo "$(shell curl -s https://icanhazdadjoke.com/)"

# Color codes used in messages.
color  := $(shell tput bold; tput setaf 6)
bright := $(shell tput bold; tput setaf 15)
dim    := $(shell tput setaf 66)
link   := $(shell tput setaf 111)
reset  := $(shell tput sgr0)

.PHONY: help vars doi-vars report watch lint clean joke

.SILENT: clean vars
