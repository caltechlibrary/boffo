# Guidelines for contributing to this project

Any constructive contributions &ndash; bug reports, pull requests (code or documentation), suggestions for improvements, and more &ndash; are welcome.

## Conduct

Everyone is asked to read and respect the [code of conduct](CODE_OF_CONDUCT.md) before participating in this project.

## Coordinating work

A quick way to find out what is currently in the near-term plans for this project is to look at the [GitHub issue tracker](https://github.com/caltechlibrary/boffo/issues), but the possibilities are not limited to what you see there &ndash; if you have ideas for new features and enhancements, please feel free to write them up as a new issue or contact the developers directly!

## Checking your changes

Before you push, run `make lint`. It checks `Code.js` with jshint and the Markdown files with markdownlint, fetching pinned versions of both with `npx`, so it needs nothing beyond Node.js, which `clasp` needs too. If [actionlint](https://github.com/rhysd/actionlint) is installed, for example with `brew install actionlint`, it also checks the GitHub Actions workflows; install it if you change them. Nothing runs these checks for you, on your computer or on GitHub.

The Makefile needs GNU Make 3.82 or later. The `make` that comes with macOS is older, so on a Mac either install a newer one with `brew install make` and run `gmake lint`, or run the checks directly:

```sh
npx --yes jshint@2.13.6 Code.js
npx --yes markdownlint-cli2@0.24.0 $(git ls-files '*.md')
actionlint
```

## Submitting contributions

Please feel free to contact the author directly, or even better, jump right in and use the standard GitHub approach of forking the repo and creating a pull request.  When committing code changes and submitting pull requests, please write a clear log message for your commits.
