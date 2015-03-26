# excel2csv

A Perl made fast MSExcel to CSV file converter.

Just run it, and you will have a new CSV file with the same name of the Excel one.

## Setup

Resolve all dependencies with CPAN:

```bash
sudo perl -MCPAN -e shell
o conf prerequisites_policy follow
o conf commit

install Text::Iconv
install Spreadsheet::XLSX
```
## Usage

```bash
./excel2csv File1.xlsx
./excel2csv File2.xls
```

## Note

One sheet per conversion is currently supported

