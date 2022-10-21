# Csv to Excel conversion - easy and fast

```sh
csv2xlsx 2.0
Giacomo R. <gcmrzz@gmail.com>

USAGE:
    csv2xlsx <SUBCOMMAND>

FLAGS:
    -h, --help       Prints help information
    -V, --version    Prints version information

SUBCOMMANDS:
    help        Prints this message or the help of the given subcommand(s)
    to-csv
    to-excel
```

## Convert to Csv

Example:
`csv2xlsx to-csv -i file.xlsx -o output.csv`

Usage:

```
USAGE:
    csv2xlsx to-csv [OPTIONS] -i <input>

FLAGS:
    -h, --help       Prints help information
    -V, --version    Prints version information

OPTIONS:
    -i <input>
    -o, --output <output>             [default: output.csv]
    -s, --sheet <sheet>              index of the sheet to be read if converting to csv [default: 0]
    -n, --sheet-name <sheet-name>    sheet name to read
```

## Convert to Excel

Example:
`csv2xlsx to-excel -i file.csv -o /tmp/output.xlsx`

Usage:

```
USAGE:
    csv2xlsx to-excel [OPTIONS]

FLAGS:
    -h, --help       Prints help information
    -V, --version    Prints version information

OPTIONS:
    -c, --column-size <column-size>     by default it depends on the content size - only useful if output is an excel
                                        file
    -i <input>...
    -o, --output <output>                [default: output.xlsx]
    -n, --sheet-name <sheet-name>...    sheet name to write
```
