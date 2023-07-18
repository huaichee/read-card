use calamine::{open_workbook, Xlsx, Reader};

fn main() {
    println!("Hello, world!");

    example();

}


fn example() {

    let mut excel: Xlsx<_> = open_workbook("test.xlsx").unwrap();
    if let Some(Ok(r)) = excel.worksheet_range("staff") {
        for row in r.rows() {
            println!("1={:?}, 2={:?}", row, row[0]);
        }
    }
}