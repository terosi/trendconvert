# Trendconvert
Converts Citect trend data to xlsx or csv.

## Usage
```
usage: trendconvert [-h] [-s] [-o TYPE] [-e] [-start DATE] [-stop DATE] [-f NUM] [-d] [-p NUM] file

positional arguments:
  file         Filename (TRENDFILE.HST)

optional arguments:
  -h, --help   show this help message and exit
  -s           Strip directories from HST data filenames. (If files are moved from original directory.)
  -o TYPE      Output file type (xls,csv)
  -e           Examine master header for dates in data files.
  -start DATE  Start date (YYYY-MM-DD)
  -stop DATE   End date (YYYY-MM-DD)
  -f NUM       Select file to export.
  -d           Do not discard invalid values from samples.
  -p NUM       Number of decimals shown in values. (Default: 1)
```

Works with Python 3.9.1
Might work with alot of other versions too.

Use at your own risk.
