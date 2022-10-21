use calamine::open_workbook_auto;
use calamine::Reader;
use std::collections::HashMap;
use std::path::PathBuf;
use structopt::StructOpt;
use xlsxwriter::*;

#[derive(StructOpt, Default)]
struct ToCsv {
    #[structopt(short, parse(from_os_str))]
    input: PathBuf,

    #[structopt(
        short,
        long,
        help = "index of the sheet to be read if converting to csv",
        default_value = "0"
    )]
    sheet: usize,

    #[structopt(short = "n", long, help = "sheet name to read")]
    sheet_name: Option<String>,

    #[structopt(short, long, default_value = "output.csv")]
    output: String,
}
#[derive(StructOpt, Default)]
struct ToExcel {
    #[structopt(short, parse(from_os_str), min_values = 2)]
    input: Vec<PathBuf>,

    #[structopt(short = "n", long, help = "sheet name to write")]
    sheet_name: Vec<String>,

    #[structopt(
        short,
        long,
        help = "by default it depends on the content size - only useful if output is an excel file"
    )]
    column_size: Option<usize>,

    #[structopt(short, long, default_value = "output.xlsx")]
    output: String,
}

#[derive(StructOpt)]
#[structopt(about = "the stupid content tracker")]
enum Args {
    ToCsv(ToCsv),
    ToExcel(ToExcel),
}

impl ToExcel {
    fn execute(&self) {
        // Open output file
        let workbook = Workbook::new(&self.output);

        for (path_i, path) in self.input.iter().enumerate() {
            let mut sheet = workbook
                .add_worksheet(Some(
                    self.sheet_name
                        .get(path_i)
                        .unwrap_or(&format!("Sheet{}", path_i + 1)),
                ))
                .expect("Could not add a worksheet");

            // Open input file
            let mut rdr = csv::ReaderBuilder::new()
                .flexible(true)
                .trim(csv::Trim::All)
                .has_headers(false)
                .from_path(&path)
                .expect(&format!("Could not read the input file {:?}", path));

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
            col_sizes.iter().for_each(|(col_i, size)| {
                let size = match self.column_size {
                    Some(size) => size as f64,
                    None => (*size as f64).max(5.0).min(70.0),
                };
                sheet
                    .set_column(*col_i as u16, *col_i as u16, size, None)
                    .unwrap();
            });
        }

        // Finish
        workbook.close().unwrap();
    }
}

impl ToCsv {
    fn execute(&self) {
        let mut excel = open_workbook_auto(&self.input).unwrap();

        let range = if let Some(name) = &self.sheet_name {
            excel
                .worksheet_range(name)
                .expect(&format!("Could not find sheet, {}", name))
                .expect(&format!("Could not find sheet, {}", name))
        } else {
            excel
                .worksheet_range_at(self.sheet)
                .expect(&format!("Could not find sheet at index, {}", self.sheet))
                .expect(&format!("Could not find sheet at index, {}", self.sheet))
        };
        let mut csv = csv::WriterBuilder::new()
            .flexible(true)
            .has_headers(false)
            .from_path(&self.output)
            .unwrap();
        range.rows().for_each(move |row| {
            let data: Vec<_> = row.iter().map(|cell| cell.to_string()).collect();
            csv.write_record(&data).unwrap();
        });
    }
}

fn main() {
    let opts: Args = Args::from_args();

    match opts {
        Args::ToCsv(x) => x.execute(),
        Args::ToExcel(x) => x.execute(),
    };
}

#[test]
fn test_all() {
    let command = ToExcel {
        input: vec![PathBuf::from("./test/input.csv")],
        output: "./test/output.xlsx".to_string(),
        ..Default::default()
    };
    command.execute();

    // Should obtain the original file
    let command = ToCsv {
        input: PathBuf::from("./test/output.xlsx"),
        output: "./test/output.csv".to_string(),
        ..Default::default()
    };
    command.execute();

    assert_eq!(
        file_size("./test/input.csv"),
        file_size("./test/output.csv"),
    );

    // This will create an excel with two sheets named "input" and "input-2"
    let command = ToExcel {
        input: vec![
            PathBuf::from("./test/input.csv"),
            PathBuf::from("./test/input-2.csv"),
        ],
        sheet_name: vec!["First sheet".to_string(), "Second sheet hello".to_string()],
        output: "./test/output_multi.xlsx".to_string(),
        ..Default::default()
    };
    command.execute();

    // This should extract only the first sheet
    let command = ToCsv {
        input: PathBuf::from("./test/output_multi.xlsx"),
        output: "./test/output_multi-1.csv".to_string(),
        ..Default::default()
    };
    command.execute();

    assert_eq!(
        file_size("./test/input.csv"),
        file_size("./test/output_multi-1.csv")
    );

    // This should extract only the second sheet
    let command = ToCsv {
        input: PathBuf::from("./test/output_multi.xlsx"),
        output: "./test/output_multi-2.csv".to_string(),
        sheet: 1,
        ..Default::default()
    };
    command.execute();

    assert_eq!(
        file_size("./test/input-2.csv"),
        file_size("./test/output_multi-2.csv")
    );

    // This should extract only the second sheet (again, by name)
    let command = ToCsv {
        input: PathBuf::from("./test/output_multi.xlsx"),
        output: "./test/output_multi-2.csv".to_string(),
        sheet_name: Some("Second sheet hello".to_string()),
        ..Default::default()
    };
    command.execute();

    assert_eq!(
        file_size("./test/input-2.csv"),
        file_size("./test/output_multi-2.csv")
    );

    fn file_size<T: AsRef<std::path::Path>>(path: T) -> u64 {
        std::os::unix::prelude::MetadataExt::size(
            &std::fs::metadata(path).expect("Cannot find input.csv"),
        )
    }
}
