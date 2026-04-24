use umya_spreadsheet::*;
use clap::Parser;

#[derive(Parser)]
#[command(name = "rexcell")]
#[command(about = "Process an Excel file using unique IDs")]
struct Args {
    /// Path to the Excel file
    #[arg(short, long, default_value = "data.xlsx")]
    file: String,

    /// Column to search for duplicate text
    #[arg(short = 's', long = "src-col", default_value = "C")]
    src_col: String,

    /// Column to write unique IDs into
    #[arg(short = 'd', long = "dest-col", default_value = "B")]
    dest_col: String,

    /// Overwrite the input file instead of creating a new one
    #[arg(short = 'i', long = "inplace", default_value = "false")]
    inplace: bool,

    /// reference table
    #[arg(short = 'r', long = "ref", default_value = "")]
    ref_table: String,

    /// reference table key column
    #[arg(short = 'k', long = "key-col", default_value = "B")]
    ref_col_key: String,

    /// reference table value column
    #[arg(short = 'v', long = "value-col", default_value = "C")]
    ref_col_value: String,

    /// update table
    #[arg(short = 'u', long = "upd", default_value = "")]
    upd_table: String,

    /// list the worksheets in the file
    #[arg(short = 'l', long = "list-sheets", default_value = "false")]
    list_sheets: bool,
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
            ref_map.insert(cell_value.clone(), cell_key.clone());
            // println!("Raw {} '{}:{}'", row, cell_value, cell_key);
        }
    }
    ref_map
}

fn apply_key_value_data(sheet: &mut Worksheet, ref_map: &std::collections::HashMap<String, String>, src_col: u32, dest_col: u32) -> Result<usize, String> {
    let mut applied = 0;

    for row in 1..=sheet.get_highest_row() {
        let cell_value = sheet.get_value((src_col, row));
        
        if !cell_value.is_empty() {
            if let Some(value) = ref_map.get(&cell_value) {
                sheet.get_cell_mut((dest_col, row)).set_value(value.clone());
                applied += 1;
                // println!("Row {}: '{}' -> ID {}", row, cell_value, value);
            } else {
                println!("Unable to find {} in ref_map!", cell_value);
            }
        }
    }

    if applied == 0 {
        Err("No key-value mapping was applied".to_string())
    } else {
        Ok(applied)
    }
}

fn get_worksheet_names_list(book: & Spreadsheet) -> Vec<String> {
    let sheets = book.get_sheet_collection();
    sheets.iter().map(|s| s.get_name().to_string()).collect()
}

fn get_worksheet_names_string(book: & Spreadsheet) -> String {
    get_worksheet_names_list(book).join(", ")
}

fn main() {
    let args = Args::parse();

    // Load the Excel file
    let path = std::path::Path::new(&args.file);
    let mut book: Spreadsheet = reader::xlsx::read(path).expect("Can't read the file");

    if args.list_sheets {
        println!("{}", get_worksheet_names_string(&book)); 
    }
    else {
        // Get the reference sheet
        let rtbl = book.get_sheet_by_name(&args.ref_table).expect("The reference sheet is not found");

        // Get the key-value entries from the reference table
        use std::collections::HashMap;
        let ref_map: HashMap<String, String> = get_ref_map(&rtbl, 
                                                    column_to_index(&args.ref_col_key), 
                                                    column_to_index(&args.ref_col_value));
        
        // Get the update sheet
        let utbl = book.get_sheet_by_name_mut(&args.upd_table).expect("The update sheet is not found");

        let applied = apply_key_value_data(
            utbl,
            &ref_map,
            column_to_index(&args.src_col),
            column_to_index(&args.dest_col),
        )
        .expect("No key-value mapping was applied");

        println!("Applied {} key-value mapping(s).", applied);

        // Save changes
        if args.inplace {
            let _ = writer::xlsx::write(&book, path).expect("Unable to write the file");
            println!("Done! The result is saved in '{}'", args.file);
        } else {
            let new_file = format!("{}_new.xlsx", args.file.trim_end_matches(".xlsx"));
            let _ = writer::xlsx::write(&book, std::path::Path::new(&new_file)).expect("Unable to write the file");
            println!("Done! The result is saved in '{}'", new_file);
        }        
    }
}