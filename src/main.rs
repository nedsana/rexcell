use clap::Parser;
use rexcell::common;
use rexcell::excell;
use log::{error};

#[derive(Parser, Debug)]
#[command(name = common::APP_NAME)]
#[command(about = common::APP_ABOUT)]
struct Args {
    #[arg(value_enum)]
    #[arg(short = 'c', long = common::ARG_LONG_COMMAND, help = common::COMMAND_FILE_HELP)]
    command: common::Command,


    #[arg(short = 't', long = common::ARG_LONG_TARGET_FILE, default_value = common::TGT_DEFAULT_EXCEL_FILE, help = common::TGT_FILE_HELP)]
    tgt_file: String,
    
    #[arg(short = 'u', long = common::ARG_LONG_UPDATE_SHEET, default_value = common::TGT_DEFAULT_TABLE, help = common::TGT_UPDATE_SHEET_HELP)]
    tgt_upd_table: String,

    #[arg(short = 's', long = common::ARG_LONG_SRC_COL, default_value = common::TGT_DEFAULT_SRC_COL, help = common::TGT_SRC_COL_HELP)]
    tgt_src_col: String,

    #[arg(short = 'd', long = common::ARG_LONG_DEST_COL, default_value = common::TGT_DEFAULT_DST_COL, help = common::TGT_DEST_COL_ACCUM_HELP)]
    tgt_dest_col: String,



    #[arg(short = 'r', long = common::ARG_LONG_REFERENCE_FILE, default_value = common::REF_DEFAULT_EXCEL_FILE, help = common::REF_FILE_HELP)]
    ref_file: String,

    #[arg(short = 'e', long = common::ARG_LONG_REFERENCE_SHEET, default_value = common::REF_DEFAULT_TABLE, help = common::REF_SHEET_HELP)]
    ref_table: String,

    #[arg(short = 'k', long = common::ARG_LONG_KEY_COL, default_value = common::REF_DEFAULT_DST_COL, help = common::REF_DEST_COL_HELP)]
    ref_col_key: String,

    #[arg(short = 'v', long = common::ARG_LONG_VALUE_COL, default_value = common::REF_DEFAULT_SRC_COL, help = common::REF_SRC_COL_HELP)]
    ref_col_value: String,



    #[arg(short = 'n', long = common::ARG_LONG_NEW_SHEET_NAME, default_value = common::TGT_DEFAULT_NEW_SHEET_NAME, help = common::NEW_SHEET_NAME_HELP)]
    new_sheet_name: String,

    #[arg(short = 'i', long = common::ARG_LONG_INPLACE, default_value = common::DEFAULT_BOOL_FALSE, help = common::INPLACE_HELP)]
    inplace: bool,



    #[arg(short = 'a', long = common::ARG_LONG_ANALYSIS_FILE, default_value = common::TGT_DEFAULT_ANALYSIS_FILE, help = common::ANALYSIS_FILE_HELP)]
    analysis_file: String,

    #[arg(short = 'o', long = common::ARG_LONG_ANALYSIS_TABLE, default_value = common::TGT_DEFAULT_ANALYSIS_TABLE, help = common::ANALYSIS_TABLE_HELP)]
    analysis_table: String,
}

// cargo run --bin rexcell -- -t ../../Test_Twins.xlsx -e "Ед. Цени" -u "Ф200" -k B -v C -s C -d B -i

fn main() 
{
    //Init the fern logger
    let res = fern::Dispatch::new()
        .format(|out, message, record| 
        {
            let level_str = format!("{:<5}", record.level().to_string());

            let file_line = format!("{}:{}", record.file().unwrap_or("unknown"), record.line().unwrap_or(0));

            let file_line_padded = format!("{:<32}", file_line);

            out.finish(format_args!("[{}] [{}] {}", level_str, file_line_padded, message))
        })
        .level(log::LevelFilter::Debug)
        .chain(std::io::stdout())
        .apply();

    match res 
    {
        Ok(_dispatcher) => 
        {
            let raw_args: Vec<_> = std::env::args_os().collect();

            std::panic::set_hook(Box::new(move |info| {
                eprintln!("Panic! cmdline args: {:?}", raw_args);
                eprintln!("{}", info);
            }));

            let args = Args::parse();

            // cargo run --bin rexcell -- -c cmd-filter-sheets -t ../Test_Excell.xlsx -u "Лист1,Лист2,Лист3" -s C -d E -n "Test"
            let cfg: common::Config = common::Config {
                command: args.command,
                tgt_file: args.tgt_file,
                tgt_upd_table: args.tgt_upd_table,
                tgt_src_col: args.tgt_src_col,
                tgt_dest_col: args.tgt_dest_col,
                ref_file: args.ref_file,
                ref_table: args.ref_table,
                ref_col_key: args.ref_col_key,
                ref_col_value: args.ref_col_value,
                new_sheet_name: args.new_sheet_name,
                inplace: args.inplace,
                analysis_file: args.analysis_file,
                analysis_table: args.analysis_table,
            };

            let res = excell::execute(&cfg);
            match res 
            {
                Ok(_) => 
                {
                }
                Err(err) => 
                {
                    error!("{}", err);
                }
            }
        }
        Err(err) => 
        {
            eprintln!("{} {}", common::ERROR_FAILED_TO_CREATE_LOGGER, err);
        }
    }
}


/*
EXAMPLES:
Update the file: read 'C' column from reference file ../Test_Excell_new.xlsx and update 'C' column in target file ../Test_Excell_new.xlsx
    cargo run --bin rexcell -- -c cmd-update-sheets -t ../Test_Excell_new.xlsx -s C -d B -u "Лист1,Лист2,Лист3" -r ../Test_Excell_new.xlsx -e "Unknown Items" -k B -v C -i
    cargo run --bin rexcell -- -c cmd-update-sheets -t ../Test_Excell_new.xlsx -s C -d B -u "Лист1,Лист2,Лист3" -r ../Test_Excell_new.xlsx -e "Test" -k B -v C -i
    cargo run --bin rexcell -- -c cmd-update-sheets -t ./Ref_Files/Test_Excell_new.xlsx -s C -d B -u "Лист1,Лист2,Лист3" -r ./Ref_Files/Test_Excell_new.xlsx -e "Test" -k B -v C -i

Update the file: get the cells in 'C' col from ref.table. Find this cell content in col 'C' from tgt.table and copy the content of ref.col 'F' to tgt.col 'F'
    cargo run --bin rexcell -- -c cmd-update-sheets -t ./Ref_Files/Test_Excell_new.xlsx -s C -d F -u "Лист1,Лист2,Лист3" -r ./Ref_Files/Test_Excell_new.xlsx -e "Test" -k F -v C -i 

Read target file ../Test_Excell.xlsx, extract unique entries from column C and store them in another sheet
    cargo run --bin rexcell -- -c cmd-filter-sheets -t ../Test_Excell.xlsx -u "Лист1,Лист2,Лист3" -s C -d E -n "Test"
    cargo run --bin rexcell -- -c cmd-filter-sheets -t ./Ref_Files/Test_Excell.xlsx -u "Лист1,Лист2,Лист3" -s C -d E -n "Test"

Lits the existing sheets in target file ../Test_Excell.xlsx
    cargo run --bin rexcell -- -c cmd-list-sheets -t ../Test_Excell.xlsx

Execute autoupdate:
RUST_BACKTRACE=1 cargo run --bin rexcell -- -c cmd-autocomplete-sheets -t "./Ref_Files/Test_Excell_T1.xlsx" -u "КСС_ОП1" -s "C" -d "B,F" -n "Prices" -i -a "./Ref_Files/Test_Excell_T1_Analysis.xlsx" -o "Analysis" >OUT

dev@ned-dev:~/Projects/Razni/rexcell$ ls Ref_Files/
'5.1.1. КСС1_Горун_Прил. 2.1 - Евроканал.xlsx'   Test_Excell_T1_Analysis.xlsx   Test_Range_Merged.xlsx                   'КСС етап 2(53337467) - ЕВРОКАНАЛ 2.04.2026.xlsx'
'5.1.1. КСС1_Горун_Прил. 2.1.xlsx'               Test_Excell_T1_small.xlsx      Test_Range_Multiline.xlsx                'ОП № 2-КСС Близнаци(ЕВРОКАНАЛ) ().xlsx'
 Test_Excell_T0.xlsx                             Test_Excell_ref.xlsx           Test_Twins.xlsx
 Test_Excell_T1.xlsx                             Test_Range_Basic.xlsx         'КСС ВОДОПРОВОД СЕВЛИЕВО(49892291).xlsx'
*/