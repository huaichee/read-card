use calamine::{open_workbook_auto, DataType, Range, Reader};
use std::env;
use std::path::PathBuf;
use xlsxwriter::prelude::*;
use pcsc::*;
use chrono::{DateTime, Local};

fn main() {
    // converts first argument into a csv (same name, silently overrides
    // if the file already exists

    let file = env::args()
        .nth(1)
        .expect("Please provide an excel file to convert");
    let sheet = env::args()
        .nth(2)
        .expect("Expecting a sheet name as second argument");

    let sce = PathBuf::from(file);
    match sce.extension().and_then(|s| s.to_str()) {
        Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
        _ => panic!("Expecting an excel file"),
    }

    let mut xl = open_workbook_auto(&sce).unwrap();
    let range = xl.worksheet_range(&sheet).unwrap().unwrap();

    let card_no = "wefwefwefw";

    write_workbook(&range, &card_no).unwrap();
}

fn write_workbook(range: &Range<DataType>, card_no: &str) -> Result<(), XlsxError>{
    let workbook = Workbook::new("simple1.xlsx")?;

    let mut sheet1 = workbook.add_worksheet(Some("Staff"))?;

    let mut row = 0;
    for r in range.rows() {
        for (i, c) in r.iter().enumerate() {
            let col = i as u16;

            match c {
                DataType::Empty => Ok(()),
                DataType::String(ref s)
                | DataType::DateTimeIso(ref s)
                | DataType::DurationIso(ref s) => if row == 0 { 
                    let bg_color = match col {
                        5 => FormatColor::Yellow,
                        _ => FormatColor::Cyan
                    };

                    sheet1.write_string(row, col, s, Some(&Format::new().set_bg_color(bg_color).set_border(FormatBorder::Thin))) 
                } else { 
                    sheet1.write_string(row, col, s, None)
                } ,
                DataType::Float(f) | DataType::DateTime(f) | DataType::Duration(f) => {
                    sheet1.write_number(row, col, f.to_owned(), None)
                }
                DataType::Int(int) => sheet1.write_number(row, col, int.to_owned() as f64, None),
                DataType::Error(ref e) => Ok(()),
                DataType::Bool(b) => sheet1.write_boolean(row, col, b.to_owned(), None),
            }?;

            if col == 5 && c == &DataType::Empty {
                let now: DateTime<Local> = Local::now();
                let formatted_date_time = now.format("%m/%d/%Y %r").to_string();
                
                let _ = sheet1.write_string(row, col, card_no, None); 
                let _ = sheet1.write_string(row, col + 1, &formatted_date_time, None); 
            }

        }
        row += 1;
    }

    workbook.close()?;

    Ok(())
}   