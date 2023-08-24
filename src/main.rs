use calamine::{open_workbook, DataType, Range, Reader, Xlsx};
use std::env;
use xlsxwriter::prelude::*;
use pcsc::*;
use chrono::{DateTime, Local};

fn main() {
    let current_directory = env::current_dir();

    let mut file_path = current_directory.unwrap().to_str().unwrap().to_owned();

    let file = "/smartstripe/doc/test2.xlsx";

    file_path.push_str(file);

    let mut excel: Xlsx<_> = open_workbook(file_path).unwrap();

    let range = excel.worksheet_range("data").unwrap().unwrap();

    let card_no = read_card().unwrap();

    write_workbook(&range, &card_no).unwrap();
}

fn read_card() -> Result<String, Error> {
    // let is_cepas = env::var("CEPAS").unwrap();

    // Establish a PC/SC context.
    let ctx = match Context::establish(Scope::User) {
        Ok(ctx) => ctx,
        Err(err) => {
            eprintln!("Failed to establish context: {}", err);
            std::process::exit(1);
        }
    };

    // List available readers.
    let mut readers_buf = [0; 2048];
    let mut readers = match ctx.list_readers(&mut readers_buf) {
        Ok(readers) => readers,
        Err(err) => {
            eprintln!("Failed to list readers: {}", err);
            std::process::exit(1);
        }
    };

    // Use the first reader.
    let reader = match readers.next() {
        Some(reader) => reader,
        None => {
            panic!("No readers are connected.");
        }
    };

    // Connect to the card.
    let card = match ctx.connect(reader, ShareMode::Shared, Protocols::ANY) {
        Ok(card) => card,
        Err(Error::NoSmartcard) => {
            panic!("A smartcard is not present in the reader.");
        }
        Err(err) => {
            eprintln!("Failed to connect to card: {}", err);
            std::process::exit(1);
        }
    };

    // if is_cepas == "1" {
    //     // Send an APDU command.
    //     let initialize_cepas = b"\x00\xA4\x00\x00\x02\x00\x00";

    //     let mut cepas_buf = [0; MAX_BUFFER_SIZE];
    //     let turn_on_cepas = match card.transmit(initialize_cepas, &mut cepas_buf) {
    //         Ok(rapdu) => rapdu,
    //         Err(err) => {
    //             eprintln!("Failed to transmit APDU command to card: {}", err);
    //             std::process::exit(1);
    //         }
    //     };

    //     println!("Cepas status: {:?}", turn_on_cepas);

    //     let apdu = b"\x90\x32\x03\x00\x00\x00";

    //     let mut rapdu_buf = [0; MAX_BUFFER_SIZE];
    //     let mut rapdu: Vec<_> = match card.transmit(apdu, &mut rapdu_buf) {
    //         Ok(rapdu) => rapdu.to_vec(),
    //         Err(err) => {
    //             eprintln!("Failed to transmit APDU command to card: {}", err);
    //             std::process::exit(1);
    //         }
    //     };

    //     // remove 144, 0
    //     rapdu.truncate(rapdu.len() - 2);


    //     let hex_value: Vec<_> = rapdu
    //         .iter()
    //         .map(|x| hex::encode_upper(x.to_be_bytes()))
    //         .collect();

    //     let hex_1 = &hex_value[8..16].join("");

    //     println!("CAN: {:?}", hex_1);

    //     println!("CSN: {:?}", &hex_value[17..25].join(":"));

    //     Ok(hex_1.to_string())

    // } else {
        let apdu = b"\xFF\xCA\x00\x00\x00";

        let mut rapdu_buf = [0; MAX_BUFFER_SIZE];
        let mut rapdu: Vec<_> = match card.transmit(apdu, &mut rapdu_buf) {
            Ok(rapdu) => rapdu.to_vec(),
            Err(err) => {
                eprintln!("Failed to transmit APDU command to card: {}", err);
                std::process::exit(1);
            }
        };

        rapdu.truncate(rapdu.len() - 2);

        let hex_value: Vec<_> = rapdu
            .iter()
            .map(|x| hex::encode_upper(x.to_be_bytes()))
            .collect();

        let decimal_value = i64::from_str_radix(&hex_value.join(""), 16);

        println!("Decimal Method 1: {:?}", decimal_value.unwrap());

         // hex method 1
        let hex_method_1 = hex_value.join(":");
        println!("Hex Method 1: {:?}", hex_method_1);

        // hex method 2
        let method_two: Vec<_> = hex_value
            .iter()
            .rev()
            .map(|x| x.to_string())
            .collect();

        let decimal_value_two = i64::from_str_radix(&method_two.join(""), 16);

        println!("Decimal Method 2: {:?}", decimal_value_two.unwrap());

        println!("Hex Method 2: {:?}", method_two.join(":"));

        Ok(hex_method_1)
    // }
}

fn write_workbook(range: &Range<DataType>, card_no: &str) -> Result<(), XlsxError>{
    let workbook = Workbook::new("smartstripe/doc/test2.xlsx")?;

    let mut sheet1 = workbook.add_worksheet(Some("data"))?;

    let mut need_fill = true;
    let mut to_fill_row = 0;

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
                DataType::Error(ref _e) => Ok(()),
                DataType::Bool(b) => sheet1.write_boolean(row, col, b.to_owned(), None),
            }?;

            if need_fill == true && col == 5 && c == &DataType::Empty {
                need_fill = false;
                to_fill_row = row;
            }

        }
        row += 1;
    }

    let now: DateTime<Local> = Local::now();
    let formatted_date_time = now.format("%m/%d/%Y %r").to_string();
    
    if need_fill == false {
        let _ = sheet1.write_string(to_fill_row, 5, card_no, None); 
        let _ = sheet1.write_string(to_fill_row, 6, &formatted_date_time, None); 
    } 
    
    workbook.close()?;

    if need_fill == true {
        eprintln!("No more space for the card to record.");
    } 

    Ok(())
}   