use calamine::{ Reader, Xlsx, open_workbook };

fn main() {
    let mut excel: Xlsx<_> = open_workbook("files/test.xlsx").unwrap();
    if let Some(Ok(r)) = excel.worksheet_range("Inuse") {
        for row in r.rows() {
            println!("{}", row[1]);
        }
    }
}
