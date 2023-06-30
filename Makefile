# =============================================================================
# @file    Makefile
# @date    2023-06-12
# @license Please see the file named LICENSE in the project directory
# @website https://github.com/caltechlibrary/boffo
#
# Run "make" or "make help" to get a list of commands in this makefile.
#
# Note: the parts involving DOIs in this makefile make 2 assumptions:
#  * The research data repository (RDM) is an InvenioRDM system
#  * The codemeta.json file contains a "relatedLink" pointing to the current
#    release archived in the RDM server.
# With these assumptions, we can automatically get the latest DOI for a
# release in RDM (because given any release, RDM can be queried for the latest
# one) and we don't have to hardwire URLs or identifiers in this makefile.
# =============================================================================

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

programs_needed = gh git jq jshint sed clasp
TEST := $(foreach p,$(programs_needed),\
	  $(if $(shell which $(p)),_,$(error Cannot find program "$(p)")))

# Set some basic variables.  These are quick to set; we set additional
# variables using "vars" but only when the others are needed.

name	:= $(strip $(shell jq -r .name codemeta.json))
version	:= $(strip $(shell jq -r .version codemeta.json))
url	:= $(strip $(shell jq -r .url codemeta.json))
related	:= $(strip $(shell jq -r .relatedLink codemeta.json))
author	:= $(strip $(shell jq -r '.author[] | .givenName + " " + .familyName' codemeta.json))
email	:= $(strip $(shell jq -r '.author[].email' codemeta.json))
license	:= $(strip $(shell jq -r .license codemeta.json))
branch	:= $(shell git rev-parse --abbrev-ref HEAD)

# The next ones depend on the convention used (at the Caltech Library) that the
# "relatedLink" field in codemeta.json points to the archived release.

rdm_current_id	 := $(shell sed -r 's|.*/(.*)$$|\1|' <<< $(related))
rdm_host_url	 := $(shell cut -d'/' -f 1-3 <<< $(related))
rdm_versions_url := $(rdm_host_url)/api/records/$(rdm_current_id)/versions


# Print help if no command is given ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# The help scheme works by looking for lines beginning with "#:" above make
# targets in this file. Originally based on code posted to Stack Overflow
# by Richard Kiefer at https://stackoverflow.com/a/59087509/743730

#: Print a summary of available commands.
help:
	@echo This is the Makefile for $(name). Available commands:
	@echo
	@grep -B1 -E "^[a-zA-Z0-9_-]+\:([^\=]|$$)" $(MAKEFILE_LIST) \
	| grep -v -- -- \
	| sed 'N;s/\n/###/' \
	| sed -n 's/^#: \(.*\)###\(.*\):.*/$(color)\2$(reset):###\1/p' \
	| column -t -s '###'

#: Summarize how to do a release using this makefile.
instructions: vars
	@echo "$$instructions_text"

define instructions_text
Steps for doing a release:
1. Run $(color)make lint$(reset) and fix problems.
2. Update the version number in the file codemeta.json.
3. Write release notes in the file CHANGES.md.
4. Run $(color)make release$(reset); it will work and eventually open a file in your editor
   for GitHub release notes. Copying the text from CHANGES.md is okay. Save it.
5. Check that everything looks okay with the GitHub release
   at https://github.com/$(repo)/releases
6. Wait for IGA to finish running its GitHub action
   at https://github.com/$(repo)/actions
7. Run $(color)make update-doi$(reset) and check the results.
endef
export instructions_text

# Gather additional values we sometimes need ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# These variables take longer to compute, and for some actions like "make help"
# they are unnecessary and annoying to wait for.
.SILENT: vars
vars:
	$(info $(dim)Gathering data – this takes a few moments …$(reset))
	$(eval repo := $(strip $(shell gh repo view | head -1 | cut -f2 -d':')))
	$(eval latest_doi := $(shell curl -s $(rdm_versions_url) | jq -r .hits.hits[0].pids.doi.identifier))
	$(info $(dim)Gathering data – this takes a few moments … Done.$(reset))

#: Print variables set in this Makefile from various sources.
report: vars
	@$(info $(color)name$(reset)           = $(name))
	@$(info $(color)version$(reset)        = $(version))
	@$(info $(color)url$(reset)            = $(url))
	@$(info $(color)author$(reset)         = $(author))
	@$(info $(color)email$(reset)          = $(email))
	@$(info $(color)license$(reset)        = $(license))
	@$(info $(color)branch$(reset)         = $(branch))
	@$(info $(color)repo$(reset)           = $(repo))
	@$(info $(color)rdm_current_id$(reset) = $(rdm_current_id))
	@$(info $(color)latest_doi$(reset)     = $(latest_doi))


# make watch ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#: Continually watch files & upload changes to Google GAS project.
watch:
	clasp -P .clasp.json.caltech  push --watch


# make lint & make test ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#: Run the code through a linter.
lint:
	jshint Code.js


# make release ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#: Make release on GitHub & update copies in Google GAS projects.
release: | confirm-release release-on-github sync-projects print-next-steps

confirm-release:
	@read -p "Have you updated the version number? [y/N] " ans && : $${ans:=N} ;\
	if [ $${ans::1} != y ]; then \
	  echo ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
	  echo ┃ Update the version number in codemeta.json first. ┃
	  echo ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
	  exit 1
	fi

update-all: update-meta update-citation

# Note that this doesn't replace "version" in codemeta.json, because that's the
# variable from which this makefile gets its version number in the first place.
update-meta:
	$(eval date := $(shell date "+%F"))
	@sed -i .bak -e '/"softwareVersion"/ s/: ".*"/: "$(version)"/' codemeta.json
	@sed -i .bak -e '/"datePublished"/ s/: ".*"/: "$(date)"/' codemeta.json

update-citation:
	$(eval date := $(shell date "+%F"))
	@sed -i .bak -e '/^date-released:/ s/".*"/"$(date)"/' CITATION.cff
	@sed -i .bak -e '/^version:/ s/".*"/"$(version)"/' CITATION.cff

edited := codemeta.json CITATION.cff

commit-updates:
	git add $(edited)
	git diff-index --quiet HEAD $(edited) || \
	    git commit -m"chore: update stored version number" $(edited)

release-on-github: | update-all commit-updates
	$(eval tmp_file := $(shell mktemp /tmp/release-notes-$(name).XXXX))
	$(eval tag := $(shell tr -d '()' <<< "$(version)" | tr ' ' '-'))
	git push -v --all
	git push -v --tags
	@$(info ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓)
	@$(info ┃ Write release notes in the file that gets opened in your  ┃)
	@$(info ┃ editor. Close the editor to complete the release process. ┃)
	@$(info ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛)
	sleep 2
	$(EDITOR) $(tmp_file)
	gh release create $(tag) -t "Release $(version)" -F $(tmp_file)

# The reason for two separate project profiles is that we have to keep two
# separate Apps Script projects on Google. The Caltech one is the real
# development copy, but it's private (because it's inside the Library's
# G-Suite/Google Workspace), so I keep a second copy in my own (mhucka)
# Google account so that I can point people to it.
sync-projects:
	clasp -P .clasp.json.caltech  push || exit 1
	clasp -P .clasp.json.public  push || exit 1

print-next-steps: vars
	@$(info ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓)
	@$(info ┃ Next steps:                                                ┃)
	@$(info ┃ 1. Check https://github.com/$(repo)/releases )
	@$(info ┃ 2. Wait for https://github.com/$(repo)/actions to finish   ┃)
	@$(info ┃ 3. Run "make update-doi" and check the results             ┃)
	@$(info ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛)

#: Update the stored DOI in CITATION.cff & README.md.
update-doi: vars
	$(eval doi_tail := $(shell cut -f'2' -d'/' <<< $(latest_doi)))
	sed -i .bak -e '/doi:/ s|doi: .*|doi: $(latest_doi)|' CITATION.cff
	sed -i .bak -e 's|records/[a-zA-Z0-9][a-zA-Z0-9][a-zA-Z0-9][a-zA-Z0-9][a-zA-Z0-9]-[a-zA-Z0-9][a-zA-Z0-9][a-zA-Z0-9][a-zA-Z0-9][a-zA-Z0-9]|records/$(doi_tail)|g' README.md
	git add CITATION.cff README.md
	git diff-index --quiet HEAD CITATION.cff README.md || \
	  (git commit -m"chore: update DOI" CITATION.cff README.md && git push -v --all)


# Cleanup and miscellaneous directives ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#: Clean this directory of temporary and backup files.
clean: clean-release clean-other
	rm -f *.bak
	rm -f tests/*.log
	@echo ✨ Cleaned! ✨

#: Print a random joke from https://icanhazdadjoke.com/
joke:
	@echo "$(shell curl -s https://icanhazdadjoke.com/)"

# Color codes used in messages.
color := $(shell tput bold; tput setaf 6)
dim   := $(shell tput setaf 66)
reset := $(shell tput sgr0)


.PHONY: help instructions vars report release test-branch tests update-all \
	update-init update-meta update-citation update-example commit-updates \
	update-doi print-next-steps clean really-clean completely-clean \
	clean-dist really-clean-dist clean-build really-clean-build \
	clean-release clean-other joke

.SILENT: clean clean-dist clean-build clean-release clean-other really-clean \
	really-clean-dist really-clean-build completely-clean
