use umya_spreadsheet::*;
use clap::Parser;

#[derive(Parser)]
#[command(name = "rexcell")]
#[command(about = "Обработка на Excel файл с уникални ID-та")]
struct Args {
    /// Път до Excel файла
    #[arg(short, long, default_value = "data.xlsx")]
    file: String,

    /// Колона за търсене на повтарящи се текстове
    #[arg(short = 's', long = "src_col", default_value = "C")]
    src_col: String,

    /// Колона за запис на уникалните номера
    #[arg(short = 'd', long = "dest_col", default_value = "B")]
    dest_col: String,

    /// Презаписва входния файл вместо създаване на нов
    #[arg(short = 'i', long = "inplace")]
    inplace: bool,

    /// reference table
    #[arg(short = 'r', long = "ref")]
    ref_table: String,

    /// reference table key column
    #[arg(short = 'k', long = "key_col", default_value = "B")]
    ref_col_key: String,

    /// reference table value column
    #[arg(short = 'v', long = "value_col", default_value = "C")]
    ref_col_value: String,

    /// update table
    #[arg(short = 'u', long = "upd")]
    upd_table: String,
}

fn column_to_index(col: &str) -> u32 {
    let mut index = 0;
    for c in col.chars() {
        index = index * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    index
}

fn get_ref_map(sheet: &Worksheet, col_key:u32, col_value: u32) -> std::collections::HashMap<String, String> {
    let mut ref_map: std::collections::HashMap<String, String> = std::collections::HashMap::new();

    for row in 1..=sheet.get_highest_row() {
        let cell_key = sheet.get_value((col_key, row));
        let cell_value = sheet.get_value((col_value, row));
        
        //invert: use the value as key
        if !cell_value.is_empty() && !cell_key.is_empty() {
            ref_map.insert(cell_key.clone(), cell_value.clone());
            // println!("Raw {} '{}:{}'", row, cell_key, cell_value);
        }
    }
    ref_map
}

fn main() {
    let args = Args::parse();

    // Load the Excel file
    let path = std::path::Path::new(&args.file);
    let mut book = reader::xlsx::read(path).expect("Can't read the file");

    // Get the reference sheet
    let rtbl = book.get_sheet_by_name(&args.ref_table).expect("The reference sheet is not found");

    // Get the key-value entries from the reference table
    use std::collections::HashMap;
    let ref_map: HashMap<String, String> = get_ref_map(&rtbl, 
                                                column_to_index(&args.ref_col_key), 
                                                column_to_index(&args.ref_col_value));

    // read the sheet to be updated and apply the key, based on value, why not reverse???
    let column_index = column_to_index(&args.src_col);
    let output_column_index = column_to_index(&args.dest_col);
    
    // Get the update sheet
    let mut utbl = book.get_sheet_by_name(&args.upd_table).expect("The update sheet is not found");
    let max_row = utbl.get_highest_row();
    utbl.get_cell_mut((0, 0)).set_value("KUR");
    // for row in 1..=max_row {
    //     let cell_value = utbl.get_value((column_index, row));
        
    //     if !cell_value.is_empty() {
    //         if let Some(value) = ref_map.get(&cell_value) {
    //             utbl.get_cell_mut((output_column_index, row)).set_value(value.clone());

    //             println!("Ред {}: Текст '{}' -> ID {}", row, cell_value, value);
    //         }
    //     }
    // }

    // // 5. Запазване на промените
    // if args.inplace {
    //     let _ = writer::xlsx::write(&book, path).expect("Грешка при запис");
    //     println!("Готово! Резултатът е записан в '{}'", args.file);
    // } else {
    //     let new_file = format!("{}_new.xlsx", args.file.trim_end_matches(".xlsx"));
    //     let _ = writer::xlsx::write(&book, std::path::Path::new(&new_file)).expect("Грешка при запис");
    //     println!("Готово! Резултатът е записан в '{}'", new_file);
    // }
}