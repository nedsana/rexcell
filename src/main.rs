use umya_spreadsheet::*;
use clap::Parser;
use rexcell::*;

#[derive(Parser)]
#[command(name = "rexcell")]
#[command(about = "Process an Excel file using unique IDs")]
struct Args {
    /// Path to the Excel file, which will be updated
    #[arg(short = 't', long = "target-file", default_value = "data.xlsx")]
    target_file: String,

    /// Path to the Excel file, where the reference data is stored. Can be the same as the target file.
    #[arg(short = 'r', long = "reference-file", default_value = "data.xlsx")]
    reference_file: String,

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
    #[arg(short = 'e', long = "reference-sheet", default_value = "")]
    ref_table: String,

    /// reference table key column
    #[arg(short = 'k', long = "key-col", default_value = "B")]
    ref_col_key: String,

    /// reference table value column
    #[arg(short = 'v', long = "value-col", default_value = "C")]
    ref_col_value: String,

    /// update table
    #[arg(short = 'u', long = "update-sheet", default_value = "")]
    upd_table: String,

    /// list the worksheets in the file
    #[arg(short = 'l', long = "list-sheets", default_value = "false")]
    list_sheets: bool,
}

// cargo run --bin rexcell -- -t ../../Test_Twins.xlsx -e "Ед. Цени" -u "Ф200" -k B -v C -s C -d B -i

fn main() {
    let args = Args::parse();

    // Load the Excel file
    let target_path = std::path::Path::new(&args.target_file);
    let mut book: Spreadsheet = reader::xlsx::read(target_path).expect("Can't read the file");

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
            let _ = writer::xlsx::write(&book, target_path).expect("Unable to write the file");
            println!("Done! The result is saved in '{}'", args.target_file);
        } else {
            let new_file = format!("{}_new.xlsx", args.target_file.trim_end_matches(".xlsx"));
            let _ = writer::xlsx::write(&book, std::path::Path::new(&new_file)).expect("Unable to write the file");
            println!("Done! The result is saved in '{}'", new_file);
        }        
    }
}
