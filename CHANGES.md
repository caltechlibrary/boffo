# Change log for Boffo

## Version 1.3.0

This version dramatically speeds up Boffo. The approach gets data from Folio in batches of 50 records at a time, and also writes the Google sheet in blocks of 50 rows. In testing, the new version gets record at a rate of between 70--100 records/second.


## Version 1.2.0

This version speeds up data fetches from the FOLIO server. Boffo should now be roughly twice as fast, which means that compared to the previous version, it can now look up twice as many barcodes before it hits the institutional time limit on Apps Script functions (which is 1800 seconds for Caltech Library users). In testing, the maximum number of records is around 1100.

Other changes in this version:
* `make lint` now lints the HTML files too.
* Minor HTML file errors have been corrected.
* The Makefile has had a few more small additions and improvements.


## Version 1.1.0

User-visible changes in this version:
* The field named "effective shelving order" is no longer printed in the output sheet for barcode lookups, following a discussion with Caltech Library staff (who reported that they didn't find it useful).
* The documentation has been slightly updated.

Other changes in this version:
* Release tags in GitHub once again start with the letter `v`, as in `v1.2.3`.
* The Makefile has been overhauled, with internal functions implemented differently and some more automation added.


## Version 1.0.0

First complete version and first release in Google Marketplace.
