# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

This version updates how Boffo logs in to FOLIO, for FOLIO servers that accept only short-lived access tokens and for institutions that use a member tenant in a FOLIO consortium.

### Added

* Boffo renews the access token with `/authn/refresh` shortly before it expires, or when FOLIO rejects it, and then retries the request once. This works in the middle of long searches. When the refresh token expires, Boffo asks for FOLIO credentials again.

### Changed

* Boffo logs in with FOLIO’s `/authn/login-with-expiry` API. It stores the short-lived access token, the refresh token, and their expiration times in the user’s properties, and never stores the user’s FOLIO password.
* When a user has no stored FOLIO login session, Boffo fills in the credentials form with the FOLIO URL and tenant id from the add-on’s script properties, replacing any older values stored for that user. This gives existing users the new tenant id the first time they log in with this version.
* The _Clear FOLIO token_ menu item deletes all of the stored token data.
* The privacy policy describes the tokens Boffo stores.
* The change log is now named `CHANGELOG.md` and follows the [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) format.

### Fixed

* FOLIO lookups failed on servers that accept only short-lived access tokens, because Boffo kept using one long-lived token from `/authn/login`.


## [1.9.1] - 2024-03-12

There are no functional changes in this release; all the changes involve build automation and updates to things like copyright statements.

* Added `alt` text to images in Markdown files.
* Added a workflow to run [Baler](https://github.com/caltechlibrary/baler).
* Added configurations and workflows for a Markdown linter, a YAML file linter, and a JSON file linter.
* Fixed linter warnings in various files.
* Updated the DOI link in the `README.md` file to use a URL that automatically resolves to the latest archived release of Boffo in [CaltechDATA](https://data.caltech.edu), so that the file does not have to be updated upon every release.
* Add a `.github/SECURITY.md` file, like a good GitHub citizen.
* Updated `Makefile` to incorporate improvements developed while working on other software.
* Updated copyright year where appropriate.


## [1.9.0] - 2023-10-05

This version introduces a new pseudo-location called "Any" in the search-by-call-numbers command, letting the user search for a given call number or call number range at any location.


## [1.8.0] - 2023-09-21

In this version, Boffo's function to search by call numbers is (hopefully) more forgiving of errors in the placement of spaces and periods.


## [1.7.1] - 2023-09-19

This version fixes a bug in fetching results from Folio that caused Boffo to return only the first 1000 results from searches that should have yielded more.


## [1.7.0] - 2023-09-12

New in this version:

* The interface for searching by call numbers now optionally lets you enter only one value, to search for that single call number at the given location.

Changes in this version:

* The previous algorithm for searching by call numbers was completely broken. This new implementation should do the right thing.
* Previously, when searching by call numbers, if nothing was found for a given call number, Boffo assumed the call number had an error in it and printed an error message to that effect. This was wrong because failing to find a call number could be due to other causes, such as if there are no items with that call number at the given location. The new version of Boffo hopefully prints a more accurate error message.


## [1.6.1] - 2023-09-05

Changes in this version:

* Boffo failed to ask for credentials when tokens expired, and instead just reported an error. Fixed; it now checks the token and brings up the credentials dialog if a new token is needed.
* Boffo prints a couple more messages while it's working, in case very large lookups take so long that the user is left wondering whether anything is happening.


## [1.6.0] - 2023-08-30

Changes in this version:

* What counts as a "barcode" was previously too specific to the patterns of barcodes used at the Caltech Library. This is no longer the case, and Boffo now accepts anything that has at least one number in it.
* The function _Find items in call number range_ will work even if the user enters the same call number in the call number range fields. The effect is to search for that single call number at the given location. This is useful in situations where a library files certain items always under the same call number (e.g., `THESIS`) and the items are distinguished on some other basis. With this change, a user can find all items with that call number at a given location.
* The _Find items in call number range_ dialog will allow a limited kind of local call number. Specifically, it will recognize single words written in capitals and consisting of only letters. Examples: `THESIS`, `FILM`.
* The values returned by _Find items in call number range_ were not being sorted. They are now sorted according to FOLIO's shelving order sort order.
* The resizing/scaling behavior of dialog windows is hopefully improved. Previously, on different browser/OS combinations, some people were getting scroll bars inside the dialogs like the call number dialog.


## [1.5.0] - 2023-08-29

This version introduces a new command: _Find items in call number range_. When invoked from the Boffo menus, it asks the user for starting and ending call numbers at a location, searches Folio for the range of items in the range of those call numbers, and writes the output to a new sheet in the spreadsheet.


## [1.4.0] - 2023-08-02

This version introduces a new menu option, _Select record fields to show_, allowing the user to select which data fields are shown for item records retrieved from FOLIO.

Other changes in this version:

* Fixed: Boffo would experience an error if you selected a single barcode and that barcode didn't exist in Folio. It will now write an empty row instead.
* Fixed: the alignment of the barcode column was always incorrect for the last row of the results spreadsheet due to an off-by-one error in the code.


## [1.3.0] - 2023-07-17

This version dramatically speeds up Boffo. The approach gets data from Folio in batches of 50 records at a time, and also writes the Google sheet in blocks of 50 rows. In testing, the new version gets record at a rate of between 70--100 records/second.


## [1.2.0] - 2023-07-12

This version speeds up data fetches from the FOLIO server. Boffo should now be roughly twice as fast, which means that compared to the previous version, it can now look up twice as many barcodes before it hits the institutional time limit on Apps Script functions (which is 1800 seconds for Caltech Library users). In testing, the maximum number of records is around 1100.

Other changes in this version:

* `make lint` now lints the HTML files too.
* Minor HTML file errors have been corrected.
* The Makefile has had a few more small additions and improvements.


## [1.1.0] - 2023-06-30

User-visible changes in this version:

* The field named "effective shelving order" is no longer printed in the output sheet for barcode lookups, following a discussion with Caltech Library staff (who reported that they didn't find it useful).
* The documentation has been slightly updated.

Other changes in this version:

* Release tags in GitHub once again start with the letter `v`, as in `v1.2.3`.
* The Makefile has been overhauled, with internal functions implemented differently and some more automation added.


## [1.0.0] - 2023-06-29

First complete version and first release in Google Marketplace.


[Unreleased]: https://github.com/caltechlibrary/boffo/compare/v1.9.1...HEAD
[1.9.1]: https://github.com/caltechlibrary/boffo/compare/v1.9.0...v1.9.1
[1.9.0]: https://github.com/caltechlibrary/boffo/compare/v1.8.0...v1.9.0
[1.8.0]: https://github.com/caltechlibrary/boffo/compare/v1.7.1...v1.8.0
[1.7.1]: https://github.com/caltechlibrary/boffo/compare/v1.7.0...v1.7.1
[1.7.0]: https://github.com/caltechlibrary/boffo/compare/v1.6.1...v1.7.0
[1.6.1]: https://github.com/caltechlibrary/boffo/compare/v1.6.0...v1.6.1
[1.6.0]: https://github.com/caltechlibrary/boffo/compare/v1.5.0...v1.6.0
[1.5.0]: https://github.com/caltechlibrary/boffo/compare/v1.4.0...v1.5.0
[1.4.0]: https://github.com/caltechlibrary/boffo/compare/v1.3.0...v1.4.0
[1.3.0]: https://github.com/caltechlibrary/boffo/compare/v1.2.0...v1.3.0
[1.2.0]: https://github.com/caltechlibrary/boffo/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/caltechlibrary/boffo/compare/1.0.0...v1.1.0
[1.0.0]: https://github.com/caltechlibrary/boffo/releases/tag/1.0.0
