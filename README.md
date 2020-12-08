# Csv to Excel conversion - easy and fast
```sh
USAGE:
    csv2xlsx [FLAGS] [OPTIONS] --input <input>

FLAGS:
    -h, --help       Prints help information
    -t, --to-csv     convert <input> to a csv file
    -V, --version    Prints version information

OPTIONS:
    -c, --column-size <column-size>     by default it depends on the content size
                                        - only useful if output is an excel file
    -i, --input <input>                
    -o, --output <output>               [default: output.xlsx]
    -s, --sheet <sheet>                 index of the sheet to be read if
                                        converting to csv [default: 0]
    -n, --sheet-name <sheet-name>       sheet name to write (or read if
                                        converting to csv)

```
