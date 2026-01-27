## R CMD check results

0 errors | 0 warnings | 0 notes

## Test environments

- local: Windows 11, R 4.5.0
- GitHub Actions: ubuntu-latest, R release
- GitHub Actions: windows-latest, R release
- GitHub Actions: macos-latest, R release
- win-builder: R-devel

## Notes

### Suggested packages not on CRAN

The package suggests `rpxl` which is available on GitHub only
(epicentre-msf/rpxl). This package provides an alternative backend for
unlocking password-protected Excel files. The functionality gracefully
degrades when rpxl is not installed, falling back to XLConnect (which is
on CRAN) or providing an informative error message.
## Downstream dependencies

This is a new package with no downstream dependencies.
