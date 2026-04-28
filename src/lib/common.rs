pub const APP_NAME: &str = "rexcell";
pub const APP_ABOUT: &str = "Process an Excel file using unique IDs";

pub const TGT_FILE_HELP: &str = "Excel file to update";
pub const TGT_SRC_COL_HELP: &str = "Column to search for duplicate text";
pub const TGT_DEST_COL_HELP: &str = "Column to write unique IDs into";
pub const TGT_UPDATE_SHEET_HELP: &str = "Update tables(sheets). Comma-separated list.";
pub const TGT_DEFAULT_EXCEL_FILE: &str = "data.xlsx";
pub const TGT_DEFAULT_SRC_COL: &str = "C";
pub const TGT_DEFAULT_DST_COL: &str = "B";
pub const TGT_DEFAULT_TABLE: &str = "";

pub const REF_FILE_HELP: &str = "Excel file, with reference data";
pub const REF_SHEET_HELP: &str = "Reference table(sheet) name";
pub const REF_KEY_COL_HELP: &str = "Reference table key column";
pub const REF_VALUE_COL_HELP: &str = "Reference table value column";
pub const REF_DEFAULT_EXCEL_FILE: &str = "data.xlsx";
pub const REF_DEFAULT_KEY_COL: &str = "B";
pub const REF_DEFAULT_VALUE_COL: &str = "C";
pub const REF_DEFAULT_TABLE: &str = "";


pub const INPLACE_HELP: &str = "Overwrite the input file instead of creating a new one";

pub const LIST_SHEETS_HELP: &str = "List of tables(sheets) in the file";
pub const LIST_SHEETS_TO_UPDATE: &str = "List of tables(sheets) to update";


pub const ARG_LONG_TARGET_FILE: &str = "tgt-file";
pub const ARG_LONG_SRC_COL: &str = "tgt-src-col";
pub const ARG_LONG_DEST_COL: &str = "tgt-dest-col";
pub const ARG_LONG_UPDATE_SHEET: &str = "tgt-sheets";
pub const ARG_LONG_REFERENCE_FILE: &str = "ref-file";
pub const ARG_LONG_REFERENCE_SHEET: &str = "ref-sheet";
pub const ARG_LONG_KEY_COL: &str = "ref-col-key";
pub const ARG_LONG_VALUE_COL: &str = "ref-col-value";
pub const ARG_LONG_INPLACE: &str = "inplace";
pub const ARG_LONG_LIST_SHEETS: &str = "list-sheets";

pub const LABEL_FILE_BROWSER: &str = "File browser";
pub const LABEL_FILE: &str = "File:";
pub const LABEL_TARGET_TEXT_FIELD_1: &str = "Target Text field 1";
pub const LABEL_TARGET_TEXT_FIELD_2: &str = "Target Text field 2";
pub const LABEL_TARGET_TEXT_FIELD_3: &str = "Target Text field 3";
pub const LABEL_REFERENCE_TEXT_FIELD_1: &str = "Reference Text field 1";
pub const LABEL_REFERENCE_TEXT_FIELD_2: &str = "Reference Text field 2";
pub const LABEL_REFERENCE_TEXT_FIELD_3: &str = "Reference Text field 3";
pub const LABEL_EXECUTION_RESULT: &str = "Execution result";
pub const BUTTON_BROWSE: &str = "Browse";
pub const BUTTON_RUN_UPDATES: &str = "Run the updates";
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


pub const DEFAULT_BOOL_FALSE: &str = "false";

pub const XLSX_EXTENSION: &str = ".xlsx";
pub const NEW_FILE_SUFFIX: &str = "_new.xlsx";

pub const ERROR_FAILED_TO_SPAWN_REXCELL: &str = "Failed to spawn rexcell: ";
pub const ERROR_FAILED_TO_START_GUI: &str = "Failed to start GUI";
pub const ERROR_CANT_READ_FILE: &str = "Can't read the file";
pub const ERROR_REFERENCE_SHEET_NOT_FOUND: &str = "The reference sheet is not found";
pub const ERROR_UPDATE_SHEET_NOT_FOUND: &str = "The update sheet is not found";
pub const ERROR_UNABLE_TO_WRITE_FILE: &str = "Unable to write the file";
pub const ERROR_MULTIPLE_REF_SHEETS: &str = "Multiple reference sheets provided!";
pub const MESSAGE_NO_KEY_VALUE_MAPPING: &str = "No key-value mapping was applied!";
pub const MESSAGE_APPLIED_MAPPINGS: &str = "Updated {} lines in table/sheet {}!";
pub const MESSAGE_DONE_SAVED: &str = "Done! The result is saved in '{}'";
pub const NO_SHEETS_FOUND: &str = "No sheets found in the file";

pub fn formatted_applied_mappings(applied: usize) -> String {
    format!("Applied {} key-value mapping(s).", applied)
}

pub fn formatted_done_saved(path: &str) -> String {
    format!("Done! The result is saved in '{}'", path)
}
