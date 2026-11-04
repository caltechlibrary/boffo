# Releasing Boffo

This page is for Boffo’s maintainers. It describes how to release a new version of Boffo and get it to users.

A release has three parts, and GitHub Actions workflows in this repository do most of the work:

1. **GitHub release.** The _Release_ workflow updates the version number and metadata, commits them to `main`, and opens a draft GitHub release. A maintainer publishes the draft. The _InvenioRDM GitHub Archiver_ workflow then archives the release in [CaltechDATA](https://data.caltech.edu), and the _Post-release_ workflow records the new DOI.
2. **Apps Script version.** The _Apps Script_ workflow pushes the release’s files to Caltech Library’s Apps Script project for Boffo and creates a numbered version of it.
3. **Marketplace update.** A maintainer sets that version number in the Google Workspace Marketplace SDK. Users get the new version only after this step, which can’t be automated.

`codemeta.json` is the master source for the version number and other metadata. The release workflow copies them into `CITATION.cff`.


## One-time setup

### GitHub Pages

The _Docs_ workflow publishes this site with GitHub Actions. In the repository’s **Settings → Pages**, set **Source** to _GitHub Actions_.

### CaltechDATA

The archiving workflow (`iga.yml`) needs a CaltechDATA token in the secret `INVENIO_TOKEN`, either in the repository or in the `caltechlibrary` organization.

### Apps Script credentials

The Apps Script API doesn’t accept service accounts, so the _Apps Script_ workflow runs `clasp` with one person’s Google credentials. They are stored as an environment secret, so that only runs a maintainer approves can read them.

1. In **Settings → Environments**, create an environment named `apps-script`. Under **Deployment protection rules**, add yourself (and any other maintainers) as **Required reviewers**. Under **Deployment branches and tags**, choose **Selected branches and tags** and add `main`.
2. Sign in to Google with a Caltech Library account that can edit Boffo’s Apps Script project, and turn on the Google Apps Script API at <https://script.google.com/home/usersettings>.
3. Log in with `clasp`, writing the credentials to a temporary file instead of `~/.clasprc.json`:

   ```shell
   npx @google/clasp@3.4.1 --auth /tmp/boffo-clasprc.json login
   ```

4. Store the file as the environment’s secret, then delete it:

   ```shell
   gh secret set CLASPRC_JSON --env apps-script --repo caltechlibrary/boffo < /tmp/boffo-clasprc.json
   rm /tmp/boffo-clasprc.json
   ```

Be aware of what this secret is:

- **It acts as you.** It holds an OAuth refresh token for your Google account, with the scopes `clasp` asks for: your Apps Script projects, the names and metadata of your Drive files, and Google Cloud (`cloud-platform`). It isn’t limited to Boffo. Anyone who can make a workflow read the secret can act as you with those scopes. Requiring a reviewer on the `apps-script` environment, and limiting it to `main`, are what keep the secret from other branches and other people’s runs.
- **It belongs to one person.** If that person leaves Caltech Library, loses access to the Apps Script project, or changes their password, the workflow fails at `clasp push`. Another maintainer then repeats steps 2–4 with their own account.
- **Caltech Library’s Google Workspace policies can end it.** A session-length limit for Google Cloud access, or a rule that blocks `clasp`’s OAuth client, makes the token stop working without warning. The workflow then fails with an authentication error such as `invalid_grant`. Repeat steps 3 and 4.
- **To revoke it**, remove `clasp` from the third-party apps with access at <https://myaccount.google.com/permissions>, and delete the `CLASPRC_JSON` secret. `clasp logout` only deletes a local file; it doesn’t revoke anything.


## Making a release

1. Make sure everything for the release is merged into `main`, and that `make lint` passes.
2. Make sure the `[Unreleased]` section of `CHANGELOG.md` describes the release, using the [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) headings (_Added_, _Changed_, _Fixed_, and so on). This section becomes the release notes.
3. On the repository’s **Actions** tab, select **Release**, then **Run workflow**. Leave the branch set to `main`, enter the new version number (such as `2.0.0`, without a `v`), and leave **Dry run** checked. When the run finishes, read its summary, which shows the changes it would commit and the release notes.
4. Run **Release** again with **Dry run** unchecked. It commits `chore: release X.Y.Z` to `main` and opens a draft release for tag `vX.Y.Z`. The run summary links to the draft.
5. Open the draft, edit the notes if needed, and click **Publish release**. Publishing creates the tag and starts the _InvenioRDM GitHub Archiver_ workflow.
6. When archiving succeeds, the _Post-release_ workflow commits `chore: update DOI for release X.Y.Z` to `main`. Check that it ran.
7. On the **Actions** tab, select **Apps Script**, then **Run workflow**. Leave the tag empty to use the latest release. Approve the run when GitHub asks for a review of the `apps-script` deployment. The run summary gives the new Apps Script version number.
8. Open the [App Configuration page](https://console.cloud.google.com/apis/api/appsmarketplace.googleapis.com/googleapps_sdk?project=boffo-add-on) of the Google Workspace Marketplace SDK in the `boffo-add-on` Cloud project. Under the Sheets add-on, set the script version to the number from step 7, and click **Save**. Users get the new version the next time Google updates their add-ons, which can take a while.

The _Release_ workflow refuses to run if the version isn’t greater than the one in `codemeta.json`, if tag `vX.Y.Z` already exists, or if the `[Unreleased]` section of `CHANGELOG.md` is empty.


## When something goes wrong

- **The _Release_ workflow fails before committing.** Nothing has changed. Fix the problem and run it again.
- **The draft release is wrong.** A draft has no tag yet. Edit the notes in the draft, or delete the draft, revert the `chore: release X.Y.Z` commit, and start again.
- **The _Post-release_ workflow fails** because CaltechDATA hasn’t finished indexing the new record. Run **Post-release** again from the **Actions** tab.
- **The _Apps Script_ workflow fails at `clasp push`** with an authentication error. The credentials have expired or were revoked. Repeat steps 3 and 4 under _Apps Script credentials_ above.


## Running the steps locally

The workflows keep their logic in shell scripts that also run on your computer. Neither script commits or pushes anything.

- `.github/scripts/prepare-release.sh 2.0.0` makes the changes a release of version 2.0.0 would make to `codemeta.json`, `CITATION.cff`, and `CHANGELOG.md`. Use `git diff` to see them, and `git checkout -- codemeta.json CITATION.cff CHANGELOG.md` to undo them.
- `.github/scripts/update-doi.sh` sets the DOI in `CITATION.cff` and `relatedLink` in `codemeta.json` to the latest archived release in CaltechDATA.

Both take `--help`.
