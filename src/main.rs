use clap::Clap;
use xlsxwriter::*;
use std::collections::HashMap;

#[derive(Clap)]
#[clap(version = "1.0", author = "Giacomo R. <gcmrzz@gmail.com>")]
struct Opts {
    #[clap(short, long)]
    input: String,
    #[clap(short, long, default_value = "output.xlsx")]
    output: String,
    #[clap(short, long, about = "by default it depends on the content size")]
    column_size: Option<usize>,
}

fn main() {
    let opts: Opts = Opts::parse();

    // Open output file
    let workbook = Workbook::new(&opts.output);
    let mut sheet = workbook.add_worksheet(None).unwrap();

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
    col_sizes.iter().for_each(|(col_i, size)| {
        let size = match opts.column_size {
            Some(size) => size as f64,
            None => (*size as f64).max(5.0).min(70.0),
        };
        sheet.set_column(*col_i as u16, *col_i as u16, size, None).unwrap();
    });

    // Finish
    workbook.close().unwrap();
}
