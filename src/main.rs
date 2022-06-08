use calamine::Error;
use calamine::{open_workbook, Reader, Sheets};
use clap::Parser;
use std::collections::HashMap;
use std::path::Path;
use xlsxwriter::*;

#[derive(Parser)]
#[clap(version = "1.2", author = "Giacomo R. <gcmrzz@gmail.com>")]
struct Opts {
    #[clap(short, long, help = "convert <input> to a csv file")]
    to_csv: bool,
    #[clap(
        short,
        long,
        help = "index of the sheet to be read if converting to csv",
        default_value = "0"
    )]
    sheet: usize,
    #[clap(
        short = 'n',
        long,
        help = "sheet name to write (or read if converting to csv)"
    )]
    sheet_name: Option<String>,
    #[clap(short, long)]
    input: String,
    #[clap(short, long, default_value = "output.xlsx")]
    output: String,
    #[clap(
        short,
        long,
        help = "by default it depends on the content size - only useful if output is an excel file"
    )]
    column_size: Option<usize>,
}
fn string_to_static_str(s: String) -> &'static str {
    Box::leak(s.into_boxed_str())
}

fn csv_to_excel(opts: Opts) {
    // Open output file
    let workbook = Workbook::new(&opts.output);
    let mut sheet = workbook
        .add_worksheet(opts.sheet_name.clone().map(|x| string_to_static_str(x)))
        .unwrap();

    // Open input file
    let mut rdr = csv::ReaderBuilder::new()
        .flexible(true)
        .trim(csv::Trim::All)
        .has_headers(false)
        .from_path(&opts.input)
        .unwrap();

    // Read csv and set excel values
    let mut col_sizes = HashMap::new();
    for (row_i, result) in rdr.records().enumerate() {
        let record = result.expect("a CSV record");
        record.iter().enumerate().for_each(|(col_i, value)| {
            if value.len() > *col_sizes.get(&col_i).unwrap_or(&0) {
                col_sizes.insert(col_i, value.len());
            }
            if value.len() > 0 {
                let res = sheet.write_string((row_i) as u32, (col_i) as u16, value, None);
                match res {
                    Ok(data) => data, 
                    Err(e) => {
                        panic!("{:?}", e.to_string())
                    }
                }
            }
        })
    }

    // Resize columns
    col_sizes.iter().for_each(move |(col_i, size)| {
        let size = match opts.column_size {
            Some(size) => size as f64,
            None => (*size as f64).max(5.0).min(70.0),
        };
        sheet
            .set_column(*col_i as u16, *col_i as u16, size, None)
            .unwrap();
    });

    // Finish
    workbook.close().unwrap();
}

fn open_workbook_xlsx<P: AsRef<Path>>(path: P) -> Result<Sheets, Error> {
    Ok(Sheets::Xlsx(open_workbook(&path).map_err(Error::Xlsx)?))
}
fn open_workbook_xls<P: AsRef<Path>>(path: P) -> Result<Sheets, Error> {
    Ok(Sheets::Xls(open_workbook(&path).map_err(Error::Xls)?))
}
fn open_workbook_xlsb<P: AsRef<Path>>(path: P) -> Result<Sheets, Error> {
    Ok(Sheets::Xlsb(open_workbook(&path).map_err(Error::Xlsb)?))
}
fn open_workbook_ods<P: AsRef<Path>>(path: P) -> Result<Sheets, Error> {
    Ok(Sheets::Ods(open_workbook(&path).map_err(Error::Ods)?))
}
fn open_auto<P: AsRef<Path>>(path: P) -> Result<Sheets, Error> {
    if let Ok(ret) = open_workbook_xlsx(&path) {
        Ok(ret)
    } else if let Ok(ret) = open_workbook_xls(&path) {
        Ok(ret)
    } else if let Ok(ret) = open_workbook_xlsb(&path) {
        Ok(ret)
    } else if let Ok(ret) = open_workbook_ods(&path) {
        Ok(ret)
    } else {
        Err(Error::Msg("Cannot detect file format"))
    }
}

fn excel_to_csv(opts: Opts) {
    let mut excel: Sheets = open_auto(&opts.input).unwrap();
    let mut csv = csv::WriterBuilder::new()
        .flexible(true)
        .has_headers(false)
        .from_path(&opts.output)
        .unwrap();
    let range = if let Some(name) = opts.sheet_name {
        excel.worksheet_range(&name).unwrap().unwrap()
    } else {
        excel.worksheet_range_at(opts.sheet).unwrap().unwrap()
    };
    range.rows().for_each(move |row| {
        let data: Vec<_> = row.iter().map(|cell| cell.to_string()).collect();
        csv.write_record(&data).unwrap();
    });
}

fn _excel_to_csv_in_ram(opts: Opts) {
    let mut excel: Sheets = open_auto(&opts.input).unwrap();
    let range = if let Some(name) = opts.sheet_name {
        excel.worksheet_range(&name).unwrap().unwrap()
    } else {
        excel.worksheet_range_at(opts.sheet).unwrap().unwrap()
    };
    let data: Vec<Vec<String>> = range
        .rows()
        .map(move |row| row.iter().map(|cell| cell.to_string()).collect())
        .collect();
    drop(excel);
    drop(range);
    let mut csv = csv::WriterBuilder::new()
        .flexible(true)
        .has_headers(false)
        .from_path(&opts.output)
        .unwrap();
    data.iter().for_each(|row| csv.write_record(row).unwrap());
}

fn main() {
    let opts: Opts = Opts::parse();

    match opts.to_csv {
        true => excel_to_csv(opts),
        false => csv_to_excel(opts),
    };
}
