use clap::{ValueEnum};
use std::fmt;

#[derive(Copy, Clone, PartialEq, Eq, PartialOrd, Ord, ValueEnum, Debug)]
pub enum Command {
    CmdListSheets,
    CmdFilterSheets,
    CmdUpdateSheets,
    CmdAutocompleteSheets,
    CmdUndefined,
}

pub const APP_NAME: &str = "rexcell";
pub const APP_ABOUT: &str = "Process an Excel file using unique IDs";

pub const COMMAND_DEFAULT: Command = Command::CmdFilterSheets;
pub const COMMAND_FILE_HELP: &str = "Command to execute.";

pub const TGT_FILE_HELP: &str = "Excel file to update";
pub const TGT_SRC_COL_HELP: &str = "Column to filter on";
pub const TGT_DEST_COL_COPY_HELP: &str = "Columns to copy to";
pub const TGT_DEST_COL_ACCUM_HELP: &str = "Columns to accumulate, on filter match";
pub const TGT_DEST_COL_HELP: &str = "Columns in the table, we copy data to.";
pub const TGT_DEST_CALCS_HELP: &str = "Calculations to apply.";
pub const TGT_UPDATE_SHEET_HELP: &str = "Update tables. Comma-separated list.";
pub const NEW_SHEET_NAME_HELP: &str = "Name of the new sheet";
pub const ANALYSIS_FILE_HELP: &str = "Excel file with analysis data";
pub const ANALYSIS_TABLE_HELP: &str = "Table with analysis data";
pub const ANALYSIS_COL_SRCH_HELP: &str = "Column with 'Begin' pattern.";
pub const ANALYSIS_COL_TERM_HELP: &str = "Column with 'End' pattern.";
pub const ANALYSIS_COLS_CP_SRC_HELP: &str = "Columns to copy from.";
pub const ANALYSIS_SRCH_PAT_HELP: &str = "'Begin' pattern";
pub const ANALYSIS_TERM_PAT_HELP: &str = "'End' pattern";

pub const TGT_DEFAULT_EXCEL_FILE: &str = "";
pub const TGT_DEFAULT_SRC_COL: &str = "C";
pub const TGT_DEFAULT_DST_COL: &str = "B,F";
pub const TGT_DEFAULT_ACC_COL: &str = "E";
pub const TGT_DEFAULT_TABLE: &str = "";
pub const TGT_DEFAULT_NEW_SHEET_NAME: &str = "Prices";
pub const TGT_DEFAULT_ANALYSIS_FILE: &str = "";
pub const TGT_DEFAULT_ANALYSIS_TABLE: &str = "";
pub const TGT_DEFAULT_CALCS: &str = "G=E*F";

pub const REF_FILE_HELP: &str = "Excel file, with reference data";
pub const REF_SHEET_HELP: &str = "Reference table name";
pub const REF_SRC_COL_HELP: &str = "Column to look for";
pub const REF_DEST_COL_HELP: &str = "Column to get data from";
pub const REF_DEFAULT_EXCEL_FILE: &str = "";
pub const REF_DEFAULT_SRC_COL: &str = "C";
pub const REF_DEFAULT_DST_COL: &str = "B,F";
pub const REF_DEFAULT_TABLE: &str = "";

pub const ANA_DEFAULT_EXCEL_FILE:  &str = "";
pub const ANA_DEFAULT_SHEET:       &str = "";
pub const ANA_DEFAULT_COL_SRCH:    &str = "A";
pub const ANA_DEFAULT_COL_TERM:    &str = "B";
pub const ANA_DEFAULT_COLS_CP_SRC: &str = "A,G";
pub const ANA_DEFAULT_SRCH_PAT:    &str = "Позиция: *, *Основание:(.*)";
pub const ANA_DEFAULT_TERM_PAT:    &str = "Общо";

pub const INPLACE_HELP: &str = "Overwrite the input file instead of creating a new one";

pub const LIST_SHEETS_HELP: &str = "List of tables in the file";
pub const LIST_SHEETS_TO_UPDATE: &str = "List of tables to update";
pub const FILTERED_SHEET: &str = "Filtered content from sheet";

pub const ARG_LONG_COMMAND: &str = "command";
pub const ARG_LONG_TARGET_FILE: &str = "tgt-file";
pub const ARG_LONG_SRC_COL: &str = "tgt-src-col";
pub const ARG_LONG_DEST_COL: &str = "tgt-dest-col";
pub const ARG_LONG_ACCUM_COL: &str = "tgt-accum-col";
pub const ARG_LONG_UPDATE_SHEET: &str = "tgt-sheets";
pub const ARG_LONG_REFERENCE_FILE: &str = "ref-file";
pub const ARG_LONG_REFERENCE_SHEET: &str = "ref-sheet";
pub const ARG_LONG_KEY_COL: &str = "ref-col-key";
pub const ARG_LONG_VALUE_COL: &str = "ref-col-value";
pub const ARG_LONG_INPLACE: &str = "inplace";
pub const ARG_LONG_LIST_SHEETS: &str = "list-sheets";
pub const ARG_LONG_NEW_SHEET_NAME: &str = "new-sheet-name";
pub const ARG_LONG_ANALYSIS_FILE: &str = "anlysis-file";
pub const ARG_LONG_ANALYSIS_TABLE: &str = "anlysis-table";

pub const LABEL_FILE_BROWSER: &str = "File browser";
pub const LABEL_FILE: &str = "File:";
pub const LABEL_TARGET_TEXT_FIELD_1: &str = "Target Text field 1";
pub const LABEL_TARGET_TEXT_FIELD_2: &str = "Target Text field 2";
pub const LABEL_TARGET_TEXT_FIELD_3: &str = "Target Text field 3";
pub const LABEL_REFERENCE_TEXT_FIELD_1: &str = "Reference Text field 1";
pub const LABEL_REFERENCE_TEXT_FIELD_2: &str = "Reference Text field 2";
pub const LABEL_REFERENCE_TEXT_FIELD_3: &str = "Reference Text field 3";
pub const LABEL_EXECUTION_RESULT: &str = "Execution result";
pub const LABEL_NEW_SHEET: &str = "Unknown Items";
pub const BUTTON_BROWSE: &str = "Browse";
pub const BUTTON_RUN_UNDEFINED: &str = "Undefined";
pub const BUTTON_RUN_UPDATES: &str = "Update File";
pub const BUTTON_FILTER_DATA: &str = "Filter Data";
pub const BUTTON_APPLY_ANALYSIS: &str = "Apply Analysis";
pub const WINDOW_TITLE: &str = "rexcell GUI";
pub const PANEL_DESCRIPTION: &str = "The top section has two identical panels.";

pub const CMD_PATH: &str = "target/debug/rexcell";
pub const CMD_ARG_TARGET: &str = "-t";
pub const CMD_ARG_REFERENCE: &str = "-r";
pub const CMD_ARG_SRC: &str = "-s";
pub const CMD_ARG_DEST: &str = "-d";
pub const CMD_ARG_UPDATE: &str = "-u";
pub const CMD_ARG_REFERENCE_SHEET: &str = "-e";
pub const CMD_ARG_KEY: &str = "-k";
pub const CMD_ARG_VALUE: &str = "-v";
pub const CMD_ARG_INPLACE: &str = "-i";

pub const TAB_LABEL_AUTOCOMPLETE: &str = "Autocomplete Tables";
pub const TAB_LABEL_FILTER: &str = "Filter Tables [manual edit]";
pub const TAB_LABEL_UPDATE: &str = "Update Tables [manual edit]";

pub const DEFAULT_BOOL_FALSE: &str = "false";

pub const XLSX_EXTENSION: &str = ".xlsx";
pub const NEW_FILE_SUFFIX: &str = "_new.xlsx";

pub const ERROR_FAILED_TO_SPAWN_REXCELL: &str = "Failed to spawn rexcell: ";
pub const ERROR_FAILED_TO_START_GUI: &str = "Failed to start GUI";
pub const ERROR_CANT_READ_FILE: &str = "Can't read file";
pub const ERROR_CANT_READ_REF_FILE: &str = "Can't read reference file";
pub const ERROR_CANT_READ_TGT_FILE: &str = "Can't read target file";
pub const ERROR_CANT_READ_ANALYSIS_FILE: &str = "Can't read analysis file";
pub const ERROR_REFERENCE_SHEET_NOT_FOUND: &str = "The reference sheet is not found";
pub const ERROR_UPDATE_SHEET_NOT_FOUND: &str = "The update sheet is not found";
pub const ERROR_ANALYSIS_SHEET_NOT_FOUND: &str = "The analysis sheet is not found";
pub const ERROR_UNABLE_TO_WRITE_FILE: &str = "Unable to write the file";
pub const ERROR_MULTIPLE_REF_SHEETS: &str = "Multiple reference sheets provided!";
pub const ERROR_NO_ROWS_UPDATED: &str = "No rows updated!";
pub const ERROR_FAILED_TO_CREATE_SHEET: &str = "Failed to create new sheet!";
pub const ERROR_FAILED_TO_ADD_SHEET: &str = "Failed to add new sheet!";
pub const ERROR_FAILED_FILTER_SHEET: &str = "Failed to filter sheet!";
pub const ERROR_INVALID_COMMAND: &str = "Invalid command!";
pub const MESSAGE_NO_KEY_VALUE_MAPPING: &str = "No key-value mapping was applied!";
pub const MESSAGE_APPLIED_MAPPINGS: &str = "Updated {} lines in table/sheet {}!";
pub const MESSAGE_DONE_SAVED: &str = "Done! The result is saved in '{}'";
pub const NO_SHEETS_FOUND: &str = "No sheets found in the file";
pub const ERROR_DEST_COL_NOT_DEFINED: &str = "Columns to update are not defined";
pub const ERROR_FAILED_TO_CREATE_LOGGER: &str = "Failed to create logger!";

// pub const REGEX_MULTILINE: &str = "-";
pub const REGEX_MULTILINE: &str = "^[ \t]*-[ \t]*(.*)$";

//to do: make these constants configurable
pub const MAX_COL: u32 = 8;
pub const MAX_ROW: u32 = 1000;

#[derive(Debug, Clone)]
pub struct Config 
{
    pub command: Command,

    pub tgt_file:               String,
    pub tgt_upd_table:          String,
    pub tgt_src_col:            String,
    pub tgt_acc_col:            String,
    pub tgt_dest_col:           String,
    pub tgt_calcs:              String,
    
    pub ref_file:               String,
    pub ref_table:              String,
    pub ref_col_key:            String,
    pub ref_col_value:          String,
    
    pub new_sheet_name:         String,
    pub inplace:                bool,

    pub analysis_file:          String,
    pub analysis_table:         String,
    pub analysis_col_srch:      String,
    pub analysis_col_term:      String,
    pub analysis_cols_cp_src:   String,
    pub analysis_srch_pat:      String,
    pub analysis_term_pat:      String,
}

impl fmt::Display for Config 
{
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result 
    {
        write!(
            f,
            "Config {{\n\
                command:              {:?}\n\
             [Filter Section]\n\
                tgt_file:             {}\n\
                tgt_upd_table:        {}\n\
                tgt_src_col:          {}\n\
                tgt_acc_col:          {}\n\
                tgt_dest_col:         {}\n\
                tgt_calcs:            {}\n\
             [Reference Section]\n\
                ref_file:             {}\n\
                ref_table:            {}\n\
                ref_col_key:          {}\n\
                ref_col_value:        {}\n\
             [Analysis Section]\n\
                analysis_file:        {}\n\
                analysis_table:       {}\n\
                analysis_col_srch:    {}\n\
                analysis_col_term:    {}\n\
                analysis_cols_cp_src: {}\n\
                analysis_srch_pat:   '{}'\n\
                analysis_term_pat:   '{}'\n\
             [Options]\n\
                new_sheet_name:       {}\n\
                inplace:              {}
            }}",
            self.command,
            self.tgt_file,
            self.tgt_upd_table,
            self.tgt_src_col,
            self.tgt_acc_col,
            self.tgt_dest_col,
            self.tgt_calcs,
            self.ref_file,
            self.ref_table,
            self.ref_col_key,
            self.ref_col_value,
            self.analysis_file,
            self.analysis_table,
            self.analysis_col_srch,
            self.analysis_col_term,
            self.analysis_cols_cp_src,
            self.analysis_srch_pat,
            self.analysis_term_pat,
            self.new_sheet_name,
            if self.inplace { "Yes" } else { "No" }
        )
    }
}

pub fn formatted_applied_mappings(applied: usize) -> String {
    format!("Applied {} key-value mapping(s).", applied)
}

pub fn formatted_done_saved(path: &str) -> String {
    format!("Done! The result is saved in '{}'", path)
}
