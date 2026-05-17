use eframe::{egui, NativeOptions};
use rfd::FileDialog;
// use std::process::Command;
use rexcell::common;
use rexcell::excell;

struct TargetData {
    path: String,
    update_sheets: String,
    src_col: String,
    dest_col: String,
    new_sheet_name: String,
}

impl Default for TargetData {
    fn default() -> Self {
        Self {
            path: String::from(common::TGT_DEFAULT_EXCEL_FILE),
            update_sheets: String::from(common::TGT_DEFAULT_TABLE),
            src_col: String::from(common::TGT_DEFAULT_SRC_COL),
            dest_col: String::from(common::TGT_DEFAULT_DST_COL),
            new_sheet_name: String::from(common::TGT_DEFAULT_NEW_SHEET_NAME),
        }
    }
}

impl TargetData {
    pub fn new(p_path: String, p_update_sheets: String, p_src_col: String, p_dest_col: String, p_new_sheet_name: String) -> Self {
        Self { 
            path: String::from(p_path),
            update_sheets: String::from(p_update_sheets),
            src_col: String::from(p_src_col),
            dest_col: String::from(p_dest_col),
            new_sheet_name: String::from(p_new_sheet_name),
        }
    }
}

struct ReferencesData { 
    path: String,
    reference_sheet: String,
    col_key: String,
    col_value: String,
}

impl Default for ReferencesData {
    fn default() -> Self {
        Self {
            path: String::from(common::REF_DEFAULT_EXCEL_FILE),
            reference_sheet: String::from(common::REF_DEFAULT_TABLE),
            col_key: String::from(common::REF_DEFAULT_KEY_COL),
            col_value: String::from(common::REF_DEFAULT_VALUE_COL),
        }
    }
}

impl ReferencesData {
    pub fn new(p_path: String, p_reference_sheet: String, p_col_key: String, p_col_value: String) -> Self {
        Self { 
            path: String::from(p_path),
            reference_sheet: String::from(p_reference_sheet),
            col_key: String::from(p_col_key),
            col_value: String::from(p_col_value),
        }
    }
}

#[derive(PartialEq)]
enum Tab 
{
    Filter,
    Update,
}

struct GuiApp 
{
    cfg_filter: TargetData,

    cfg_update_tgt: TargetData,
    cfg_update_ref: ReferencesData,

    output_text: String,
    error: String,

    active_tab: Tab,
}

impl Default for GuiApp 
{
    fn default() -> Self 
    {
        Self 
        {
            cfg_filter: TargetData::new( common::TGT_DEFAULT_EXCEL_FILE.to_string(), 
                                common::TGT_DEFAULT_TABLE.to_string(), 
                                      common::TGT_DEFAULT_SRC_COL.to_string(), 
                                      common::TGT_DEFAULT_ACC_COL.to_string(), 
                                      common::TGT_DEFAULT_NEW_SHEET_NAME.to_string()),

            cfg_update_tgt: TargetData::new( common::TGT_DEFAULT_EXCEL_FILE.to_string(), 
                                common::TGT_DEFAULT_TABLE.to_string(), 
                                    common::TGT_DEFAULT_SRC_COL.to_string(), 
                                    common::TGT_DEFAULT_DST_COL.to_string(), 
                                    common::TGT_DEFAULT_NEW_SHEET_NAME.to_string()),

            cfg_update_ref: ReferencesData::new( common::REF_DEFAULT_EXCEL_FILE.to_string(), 
                                    common::REF_DEFAULT_TABLE.to_string(), 
                                    common::REF_DEFAULT_KEY_COL.to_string(), 
                                    common::REF_DEFAULT_VALUE_COL.to_string()),

            output_text: String::new(),
            error: String::new(),

            active_tab: Tab::Filter,
        }
    }
}


impl GuiApp 
{
    // fn get_sheets_list(&mut self, file_path: &str) -> Result<String, String> 
    fn get_sheets_list(file_path: &str) -> Result<String, String> 
    {
        let result = excell::get_worksheet_names(std::path::Path::new(&file_path));
        match result 
        {
            Ok(names) => {
                if names.len() > 0 
                {
                    Ok(names)
                } 
                else 
                {
                    Err(format!("{} {}", common::NO_SHEETS_FOUND, file_path))
                }
            }
            Err(err) => Err(format!("{}", err)),
        }
    }

    fn handle_result(res: &Result<(Vec<String>, Vec<String>), String>) -> (String, String)
    {
        let mut out_res = String::new();
        let mut out_err = String::new();

        match res {
            Ok(lines) => {
                for line in &lines.0 {
                    out_res.push_str(line);
                    out_res.push_str("\n");
                }
                for line in &lines.1 {
                    out_res.push_str(line);
                    out_res.push_str("\n");
                }
            }
            Err(err) => {
                out_err.push_str(err);
            }
        }

        (out_res, out_err)
    }

    fn draw_button_browse<FOnClick>(
        ui: &mut egui::Ui, 
        txt_label: &str, 
        txt_button: &str, 
        path:&mut String, 
        onClick: FOnClick
    )
    where FOnClick: FnOnce(&str)
    {
        ui.label(txt_label);
        ui.add_space(4.0);
        ui.horizontal(|ui|
        {
            ui.text_edit_singleline(path);

            if ui.button(txt_button).clicked() 
            {
                if let Some(path_buf) = FileDialog::new().pick_file() 
                {
                    if let Some(path_str) = path_buf.to_str() {
                        *path = path_str.to_string();
                        onClick(path_str);
                    }
                }
            }
        });
    }

    // fn draw_filter_section(&mut self, ui: &mut egui::Ui, cfg: &mut TargetData)
    fn draw_filter_section(ui: &mut egui::Ui, cfg: &mut TargetData, out_res: &mut String, out_err: &mut String, do_filter: bool) 
    {
        egui::Frame::group(ui.style()).show(ui, |ui| 
        {
            Self::draw_button_browse(ui, common::TGT_FILE_HELP, common::BUTTON_BROWSE, &mut cfg.path,
                |path_str| {
                    Self::get_sheets_list(path_str)
                        .map(|sheets| cfg.update_sheets = sheets)
                        .map_err(|err| *out_err = err)
                        .ok();
                },
            );

            ui.add_space(8.0);
            ui.label(common::LIST_SHEETS_TO_UPDATE);
            ui.text_edit_singleline(&mut cfg.update_sheets);

            ui.add_space(4.0);
            ui.label(common::TGT_SRC_COL_HELP);
            ui.text_edit_singleline(&mut cfg.src_col);

            ui.add_space(4.0);
            ui.label(common::TGT_DEST_COL_HELP);
            ui.text_edit_singleline(&mut cfg.dest_col);
            
            ui.add_space(4.0);
            if do_filter
            {
                ui.add_space(4.0);
                ui.label(common::NEW_SHEET_NAME_HELP);
                ui.text_edit_singleline(&mut cfg.new_sheet_name);

                if ui.button(common::BUTTON_FILTER_DATA).clicked()
                {
                    // cargo run --bin rexcell -- -c cmd-filter-sheets -t ../Test_Excell.xlsx -u "Лист1,Лист2,Лист3" -s C -d E -n "Test"
                    let cfg: common::Config = common::Config {
                        command: common::Command::CmdFilterSheets,
                        tgt_file: cfg.path.clone(), 
                        tgt_upd_table: cfg.update_sheets.clone(),
                        tgt_src_col: cfg.src_col.clone(),
                        tgt_dest_col: cfg.dest_col.clone(),
                        ref_file: "".to_string(),
                        ref_table: "".to_string(),
                        ref_col_key: "".to_string(),
                        ref_col_value: "".to_string(),
                        new_sheet_name: cfg.new_sheet_name.clone(),
                        inplace: true,
                    };

                    let res = excell::execute(&cfg);

                    let out = Self::handle_result(&res);

                    if 0 < out.1.len() //error found
                    {
                        out_res.clear();
                        *out_err = format!("Failed to filter file {}!\n{}\n", cfg.tgt_file, out.1);
                    }
                    else //ok
                    {
                        out_err.clear();
                        *out_res = if cfg.inplace { format!("Filtered file {}!\n{}\n", cfg.tgt_file, out.0) } 
                                   else { 
                                        let new_file = format!("{}{}", cfg.tgt_file.trim_end_matches(common::XLSX_EXTENSION), common::NEW_FILE_SUFFIX);
                                        format!("Filtered to file {}! {}\n", new_file, out.0) };
                    }
                }
            }
        });
    }

    // fn draw_cfg_update_ref(&mut self, ui: &mut egui::Ui) 
    fn draw_cfg_update_ref(ui: &mut egui::Ui, tgt_cfg: &mut TargetData, ref_cfg: &mut ReferencesData, out_res: &mut String, out_err: &mut String) 
    {
        egui::Frame::group(ui.style()).show(ui, |ui| 
        {
            ui.label(common::REF_FILE_HELP);
            ui.add_space(4.0);
            ui.horizontal(|ui| {
                ui.label(common::LABEL_FILE);
                ui.text_edit_singleline(&mut ref_cfg.path);
                if ui.button(common::BUTTON_BROWSE).clicked() {
                    if let Some(path_buf) = FileDialog::new().pick_file() {
                        if let Some(path_str) = path_buf.to_str() {
                            ref_cfg.path = path_str.to_string();
                            Self::get_sheets_list(path_str)
                                .map(|sheets| ref_cfg.reference_sheet = sheets)
                                .map_err(|err| *out_err = err)
                                .ok();
                        }
                    }
                }
            });

            ui.add_space(8.0);
            ui.label(common::REF_SHEET_HELP);
            ui.text_edit_singleline(&mut ref_cfg.reference_sheet);

            ui.add_space(4.0);
            ui.label(common::REF_KEY_COL_HELP);
            ui.text_edit_singleline(&mut ref_cfg.col_key);

            ui.add_space(4.0);
            ui.label(common::REF_VALUE_COL_HELP);
            ui.text_edit_singleline(&mut ref_cfg.col_value);

            ui.add_space(4.0);
            if ui.button(common::BUTTON_RUN_UPDATES).clicked()
            {
                let ref_sheets: Vec<String> = ref_cfg.reference_sheet.split(',').map(str::trim).map(String::from).collect();
                
                if 1 == ref_sheets.len() 
                {
                    // cargo run --bin rexcell -- -c cmd-update-sheets -t ../Test_Excell_new.xlsx -s C -d B -u "Лист1,Лист2,Лист3" -r ../Test_Excell_new.xlsx -e "Test" -k B -v C -i
                    let cfg: common::Config = common::Config {
                        command: common::Command::CmdUpdateSheets,
                        tgt_file: tgt_cfg.path.clone(), 
                        tgt_upd_table: tgt_cfg.update_sheets.clone(),
                        tgt_src_col: tgt_cfg.src_col.clone(),
                        tgt_dest_col: tgt_cfg.dest_col.clone(),
                        ref_file: ref_cfg.path.clone(),
                        ref_table: ref_cfg.reference_sheet.clone(),
                        ref_col_key: ref_cfg.col_key.clone(),
                        ref_col_value: ref_cfg.col_value.clone(),
                        new_sheet_name: tgt_cfg.new_sheet_name.clone(),
                        inplace: true,
                    };

                    let res = excell::execute(&cfg);

                    let out = Self::handle_result(&res);

                    if 0 < out.1.len() //error found
                    {
                        out_res.clear();
                        *out_err = format!("Failed to update file {}! {}\n", cfg.tgt_file, out.1);
                    }
                    else //ok
                    {
                        out_err.clear();
                        *out_res = if cfg.inplace { format!("Updated file {}! {}\n", cfg.tgt_file, out.0) } 
                                   else { 
                                        let new_file = format!("{}{}", cfg.tgt_file.trim_end_matches(common::XLSX_EXTENSION), common::NEW_FILE_SUFFIX);
                                        format!("Updated to file {}! {}\n", new_file, out.0) };
                    }
                }
                else
                {
                    *out_err = String::from(common::ERROR_MULTIPLE_REF_SHEETS);
                }
            }
        });
    }
}

impl eframe::App for GuiApp 
{
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) 
    {
        egui::CentralPanel::default().show(ctx, |ui| 
        {
            ui.vertical(|ui| 
            {
                ui.heading(common::WINDOW_TITLE);
                // ui.label(common::PANEL_DESCRIPTION);

                ui.add_space(8.0);

                ui.horizontal(|ui| 
                {
                    ui.selectable_value(&mut self.active_tab, Tab::Filter, common::TAB_LABEL_FILTER);
                    ui.selectable_value(&mut self.active_tab, Tab::Update, common::TAB_LABEL_UPDATE);
                });

                match self.active_tab 
                {
                    Tab::Filter => 
                    {
                        egui::Frame::group(ui.style()).show(ui, |ui| 
                            {
                                ui.columns(2, |columns| 
                                {
                                    Self::draw_filter_section(&mut columns[0], &mut self.cfg_filter, &mut self.output_text, &mut self.error, true);
                                });
                            });
                    }
                    
                    Tab::Update => 
                    {
                        egui::Frame::group(ui.style()).show(ui, |ui| 
                        {
                            ui.columns(2, |columns| 
                            {
                                Self::draw_filter_section(&mut columns[0], &mut self.cfg_update_tgt, &mut self.output_text, &mut self.error, false);
                                Self::draw_cfg_update_ref(&mut columns[1], &mut self.cfg_update_tgt, &mut self.cfg_update_ref, &mut self.output_text, &mut self.error);
                            });
                        });
                    }
                }

                ui.add_space(12.0);

                egui::Frame::group(ui.style()).show(ui, |ui| {
                    ui.label(common::LABEL_EXECUTION_RESULT);
                    ui.add_space(4.0);

                    egui::ScrollArea::vertical()
                        .id_source("execution_result_scroll") 
                        .max_height(400.0) 
                        .auto_shrink([false; 2]) 
                        .show(ui, |ui| {
                            ui.add(
                                egui::TextEdit::multiline(&mut self.output_text)
                                    .desired_rows(16)
                                    .desired_width(f32::INFINITY)
                                    .interactive(false) 
                                    .lock_focus(true),
                            );
                        });
                });

                if !self.error.is_empty() {
                    ui.add_space(8.0);
                    ui.colored_label(egui::Color32::RED, &self.error);
                }
            });
        });
    }

}

fn main() {
    let options = NativeOptions::default();
    eframe::run_native(common::WINDOW_TITLE, options, 
        Box::new(|_cc| Box::new(GuiApp::default()))).expect(common::ERROR_FAILED_TO_START_GUI);
}

// cargo run --bin gui
// cargo run --bin rexcell -- -t ../../Test_Twins.xlsx -e "Ед. Цени" -u "Ф200" -k B -v C -s C -d B -i