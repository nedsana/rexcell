use umya_spreadsheet::*;
use clap::Parser;
use rexcell::*;
use rexcell::common;

#[derive(Parser, Debug)]
#[command(name = common::APP_NAME)]
#[command(about = common::APP_ABOUT)]
struct Args {
    #[arg(short = 't', long = common::ARG_LONG_TARGET_FILE, default_value = common::TGT_DEFAULT_EXCEL_FILE, help = common::TGT_FILE_HELP)]
    target_file: String,
    
    #[arg(short = 'u', long = common::ARG_LONG_UPDATE_SHEET, default_value = common::TGT_DEFAULT_TABLE, help = common::TGT_UPDATE_SHEET_HELP)]
    upd_table: String,

    #[arg(short = 's', long = common::ARG_LONG_SRC_COL, default_value = common::TGT_DEFAULT_SRC_COL, help = common::TGT_SRC_COL_HELP)]
    src_col: String,

    #[arg(short = 'd', long = common::ARG_LONG_DEST_COL, default_value = common::TGT_DEFAULT_DST_COL, help = common::TGT_DEST_COL_HELP)]
    dest_col: String,



    #[arg(short = 'r', long = common::ARG_LONG_REFERENCE_FILE, default_value = common::REF_DEFAULT_EXCEL_FILE, help = common::REF_FILE_HELP)]
    reference_file: String,

    #[arg(short = 'e', long = common::ARG_LONG_REFERENCE_SHEET, default_value = common::REF_DEFAULT_TABLE, help = common::REF_SHEET_HELP)]
    ref_table: String,

    #[arg(short = 'k', long = common::ARG_LONG_KEY_COL, default_value = common::REF_DEFAULT_KEY_COL, help = common::REF_KEY_COL_HELP)]
    ref_col_key: String,

    #[arg(short = 'v', long = common::ARG_LONG_VALUE_COL, default_value = common::REF_DEFAULT_VALUE_COL, help = common::REF_VALUE_COL_HELP)]
    ref_col_value: String,



    #[arg(short = 'i', long = common::ARG_LONG_INPLACE, default_value = common::DEFAULT_BOOL_FALSE, help = common::INPLACE_HELP)]
    inplace: bool,

    #[arg(short = 'l', long = common::ARG_LONG_LIST_SHEETS, default_value = common::DEFAULT_BOOL_FALSE, help = common::LIST_SHEETS_HELP)]
    list_sheets: bool,
}

// cargo run --bin rexcell -- -t ../../Test_Twins.xlsx -e "Ед. Цени" -u "Ф200" -k B -v C -s C -d B -i

fn main() {
    let raw_args: Vec<_> = std::env::args_os().collect();

    std::panic::set_hook(Box::new(move |info| {
        eprintln!("Panic! cmdline args: {:?}", raw_args);
        eprintln!("{}", info);
    }));

    let args = Args::parse();

    if args.list_sheets {
        // Load the update Excel file
        let target_path = std::path::Path::new(&args.target_file);
        let bk: Spreadsheet = reader::xlsx::read(target_path).expect(common::ERROR_CANT_READ_FILE);
        println!("{}", get_worksheet_names_string(&bk)); 
    }
    else {
        // Load the reference Excel file
        let ref_path = std::path::Path::new(&args.reference_file);
        let rbook: Spreadsheet = reader::xlsx::read(ref_path).expect(common::ERROR_CANT_READ_FILE);

        // Load the update Excel file
        let target_path = std::path::Path::new(&args.target_file);
        let mut ubook: Spreadsheet = reader::xlsx::read(target_path).expect(common::ERROR_CANT_READ_FILE);

        // Get the reference sheet
        let rtbl = rbook.get_sheet_by_name(&args.ref_table).expect(common::ERROR_REFERENCE_SHEET_NOT_FOUND);

        // Get the key-value entries from the reference table
        use std::collections::HashMap;
        let ref_map: HashMap<String, String> = get_ref_map_by_strings(&rtbl, &args.ref_col_key, &args.ref_col_value);

        for utbln in args.upd_table.split(',') {
            // Get the update sheet
            let utbl = ubook.get_sheet_by_name_mut(&utbln).expect(common::ERROR_UPDATE_SHEET_NOT_FOUND);

            let applied = apply_key_value_data_by_strings(utbl, &ref_map, &args.src_col, &args.dest_col).expect(common::MESSAGE_NO_KEY_VALUE_MAPPING);

            println!("Updated {} lines in table/sheet '{}'!", applied, utbln);
            println!("");
        }

        // Save changes
        if args.inplace {
            let _ = writer::xlsx::write(&ubook, target_path).expect(common::ERROR_UNABLE_TO_WRITE_FILE);
            println!("{}", common::formatted_done_saved(&args.target_file));
        } else {
            let new_file = format!("{}{}", args.target_file.trim_end_matches(common::XLSX_EXTENSION), common::NEW_FILE_SUFFIX);
            let _ = writer::xlsx::write(&ubook, std::path::Path::new(&new_file)).expect(common::ERROR_UNABLE_TO_WRITE_FILE);
            println!("{}", common::formatted_done_saved(&new_file));
        }        
    }
}
