use clap::Clap;
use std::collections::HashMap;
use xlsxwriter::*;
extern crate calamine;

#[derive(Clap)]
#[clap(version = "1.1", author = "Giacomo R. <gcmrzz@gmail.com>")]
struct Opts {
    #[clap(short, long, about = "convert <input> to a csv file")]
    to_csv: bool,
    #[clap(
        short,
        long,
        about = "index of the sheet to be read if converting to csv",
        default_value = "0"
    )]
    sheet: usize,
    #[clap(short = "n", long, about = "sheet name to write (or read if converting to csv)")]
    sheet_name: Option<String>,
    #[clap(short, long)]
    input: String,
    #[clap(short, long, default_value = "output.xlsx")]
    output: String,
    #[clap(
        short,
        long,
        about = "by default it depends on the content size - only useful if output is an excel file"
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
                sheet
                    .write_string((row_i) as u32, (col_i) as u16, value, None)
                    .unwrap();
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

fn excel_to_csv(opts: Opts) {
    use calamine::*;

    let mut excel = open_workbook_auto(&opts.input).expect("Cannot open input file");
    let mut csv = csv::WriterBuilder::new()
        .flexible(true)
        .has_headers(false)
        .from_path(&opts.output)
        .unwrap();
    let range;
    if let Some(name) = opts.sheet_name {
        range = excel.worksheet_range(&name).unwrap().unwrap();
    } else {
        range = excel.worksheet_range_at(opts.sheet).unwrap().unwrap();
    }
    range.rows().for_each(move |row| {
        let data: Vec<_> = row.iter().map(|cell| cell.to_string()).collect();
        csv.write_record(&data).unwrap();
    });
}

fn main() {
    let opts: Opts = Opts::parse();

    match opts.to_csv {
        true => excel_to_csv(opts),
        false => csv_to_excel(opts),
    };
}
