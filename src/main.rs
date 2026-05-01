use clap::Parser;
use rexcell::common;

#[derive(Parser, Debug)]
#[command(name = common::APP_NAME)]
#[command(about = common::APP_ABOUT)]
struct Args {
    #[arg(short = 't', long = common::ARG_LONG_TARGET_FILE, default_value = common::TGT_DEFAULT_EXCEL_FILE, help = common::TGT_FILE_HELP)]
    tgt_file: String,
    
    #[arg(short = 'u', long = common::ARG_LONG_UPDATE_SHEET, default_value = common::TGT_DEFAULT_TABLE, help = common::TGT_UPDATE_SHEET_HELP)]
    tgt_upd_table: String,

    #[arg(short = 's', long = common::ARG_LONG_SRC_COL, default_value = common::TGT_DEFAULT_SRC_COL, help = common::TGT_SRC_COL_HELP)]
    tgt_src_col: String,

    #[arg(short = 'd', long = common::ARG_LONG_DEST_COL, default_value = common::TGT_DEFAULT_DST_COL, help = common::TGT_DEST_COL_HELP)]
    tgt_dest_col: String,



    #[arg(short = 'r', long = common::ARG_LONG_REFERENCE_FILE, default_value = common::REF_DEFAULT_EXCEL_FILE, help = common::REF_FILE_HELP)]
    ref_file: String,

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
        println!("{}", rexcell::get_worksheet_names(std::path::Path::new(&args.tgt_file))); 
    }
    else {
        let cfg: common::Config = common::Config {
            tgt_file: args.tgt_file,
            tgt_upd_table: args.tgt_upd_table,
            tgt_src_col: args.tgt_src_col,
            tgt_dest_col: args.tgt_dest_col,
            ref_file: args.ref_file,
            ref_table: args.ref_table,
            ref_col_key: args.ref_col_key,
            ref_col_value: args.ref_col_value,
            inplace: args.inplace,
            list_sheets: args.list_sheets,
        };

        let res =rexcell::execute(&cfg);

        match res {
            Ok(lines) => {
                for line in &lines.0 {
                    println!("{}", line);
                }
                for line in &lines.1 {
                    println!("{}", line);
                }
                if cfg.inplace {
                    println!("Updated {} lines. {}", lines.0.len(), common::formatted_done_saved(&cfg.tgt_file));
                } else {
                    let new_file = format!("{}{}", cfg.tgt_file.trim_end_matches(common::XLSX_EXTENSION), common::NEW_FILE_SUFFIX);
                    println!("Updated {} lines. {}", lines.0.len(), common::formatted_done_saved(&new_file));
                }
            }
            Err(err) => {
                println!("Failed to update {}: {}", cfg.tgt_file, err);
            }
        }
    }
}
