use calamine::{open_workbook_auto, DataType, Range, Reader};
use std::env;
use std::path::PathBuf;
use xlsxwriter::prelude::*;

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

    write_workbook(&range).unwrap();
}

fn write_workbook(range: &Range<DataType>) -> Result<(), XlsxError>{
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
                    let mut bg_color = FormatColor::Cyan;
                    if col == 5 {
                        bg_color = FormatColor::Yellow;
                    }

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
            
        }
        row += 1;
    }

    workbook.close()?;

    Ok(())
}   