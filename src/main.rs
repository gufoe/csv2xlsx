use clap::Clap;
use xlsxwriter::*;

#[derive(Clap)]
#[clap(version = "1.0", author = "Giacomo R. <gcmrzz@gmail.com>")]
struct Opts {
    #[clap(short, long)]
    input: String,
    #[clap(short, long, default_value = "output.xlsx")]
    output: String,
}

fn main() {
    let opts: Opts = Opts::parse();

    let workbook = Workbook::new(&opts.output);

    let mut sheet = workbook.add_worksheet(None).unwrap();

    let mut rdr = csv::ReaderBuilder::new()
        .flexible(true)
        .trim(csv::Trim::All)
        .has_headers(false)
        .from_path(opts.input)
        .unwrap();
    for (row_i, result) in rdr.records().enumerate() {
        let record = result.expect("a CSV record");
        record.iter().enumerate().for_each(|(col_i, value)| {
            if value.len() > 0 {
                sheet
                    .write_string((row_i) as u32, (col_i) as u16, value, None)
                    .unwrap();
            }
        })
    }

    sheet.set_column(0, 1000, 20.0, None).unwrap();
    workbook.close().unwrap();
}
