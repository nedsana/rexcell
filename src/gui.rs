use eframe::{egui, NativeOptions};
use rfd::FileDialog;
// use std::process::Command;
use rexcell::common;

struct TargetData {
    path: String,
    update_sheet: String,
    src_col: String,
    dest_col: String,
    new_file_name: String,
}

impl Default for TargetData {
    fn default() -> Self {
        Self {
            path: String::from(common::TGT_DEFAULT_EXCEL_FILE),
            update_sheet: String::from(common::TGT_DEFAULT_TABLE),
            src_col: String::from(common::TGT_DEFAULT_SRC_COL),
            dest_col: String::from(common::TGT_DEFAULT_DST_COL),
            new_file_name: String::from(common::TGT_DEFAULT_NEW_SHEET_NAME),
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

#[derive(PartialEq)]
enum Tab 
{
    Filter,
    Update,
}

struct GuiApp {
    target_section: TargetData,
    reference_section: ReferencesData,
    output_text: String,
    error: String,
    active_tab: Tab,
}

impl Default for GuiApp {
    fn default() -> Self {
        Self {
            target_section: TargetData::default(),
            reference_section: ReferencesData::default(),
            output_text: String::new(),
            error: String::new(),
            active_tab: Tab::Filter,
        }
    }
}

impl GuiApp 
{
    fn get_sheets_list(&mut self, file_path: &str) -> Result<String, String> 
    {
        let result = rexcell::get_worksheet_names(std::path::Path::new(&file_path));
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

    fn draw_target_section(&mut self, ui: &mut egui::Ui) 
    {
        egui::Frame::group(ui.style()).show(ui, |ui| 
        {
            ui.label(common::TGT_FILE_HELP);
            ui.add_space(4.0);
            ui.horizontal(|ui| {
                ui.label(common::LABEL_FILE);
                ui.text_edit_singleline(&mut self.target_section.path);
                if ui.button(common::BUTTON_BROWSE).clicked() {
                    if let Some(path_buf) = FileDialog::new().pick_file() {
                        if let Some(path_str) = path_buf.to_str() {
                            self.target_section.path = path_str.to_string();
                            self.get_sheets_list(path_str)
                                .map(|sheets| self.target_section.update_sheet = sheets)
                                .map_err(|err| self.error = err)
                                .ok();
                        }
                    }
                }
            });

            ui.add_space(8.0);
            ui.label(common::LIST_SHEETS_TO_UPDATE);
            ui.text_edit_singleline(&mut self.target_section.update_sheet);

            ui.add_space(4.0);
            ui.label(common::TGT_SRC_COL_HELP);
            ui.text_edit_singleline(&mut self.target_section.src_col);

            ui.add_space(4.0);
            ui.label(common::TGT_DEST_COL_HELP);
            ui.text_edit_singleline(&mut self.target_section.dest_col);

            ui.add_space(4.0);
            ui.label(common::NEW_SHEET_NAME_HELP);
            ui.text_edit_singleline(&mut self.target_section.new_file_name);

            ui.add_space(4.0);
            if ui.button(common::BUTTON_FILTER_DATA).clicked()
            {
                self.error.clear();
                self.output_text.clear();

                let ref_sheets: Vec<String> = self.reference_section.reference_sheet.split(',').map(str::trim).map(String::from).collect();
                
                if 1 == ref_sheets.len() 
                {
                    // cargo run --bin rexcell -- -c cmd-filter-sheets -t ../Test_Excell.xlsx -u "Лист1,Лист2,Лист3" -s C -d E -n "Test"
                    let cfg: common::Config = common::Config {
                        command: common::Command::CmdFilterSheets,
                        tgt_file: self.target_section.path.clone(), 
                        tgt_upd_table: self.target_section.update_sheet.clone(),
                        tgt_src_col: self.target_section.src_col.clone(),
                        tgt_dest_col: self.target_section.dest_col.clone(),
                        ref_file: "".to_string(),
                        ref_table: "".to_string(),
                        ref_col_key: "".to_string(),
                        ref_col_value: "".to_string(),
                        new_sheet_name: self.target_section.new_file_name.clone(),
                        inplace: false,
                    };

                    let res = rexcell::execute(&cfg);

                    match res {
                        Ok(lines) => {
                            for line in &lines.0 {
                                self.output_text.push_str(line);
                                self.output_text.push_str("\n");
                            }
                            for line in &lines.1 {
                                self.output_text.push_str(line);
                                self.output_text.push_str("\n");
                            }
                            if cfg.inplace {
                                self.output_text.push_str(format!("Updated {} lines. {}\n", lines.0.len(), common::formatted_done_saved(&cfg.tgt_file)).as_str());
                            } else {
                                let new_file = format!("{}{}", cfg.tgt_file.trim_end_matches(common::XLSX_EXTENSION), common::NEW_FILE_SUFFIX);
                                self.output_text.push_str(format!("Updated {} lines. {}\n", lines.0.len(), common::formatted_done_saved(&new_file)).as_str());
                            }
                        }
                        Err(err) => {
                            self.error = format!("Failed to update {}: {}", cfg.tgt_file, err);
                        }
                    }
                }
                else
                {
                    self.error = String::from(common::ERROR_MULTIPLE_REF_SHEETS);
                }
            }
        });
    }

    fn draw_reference_section(&mut self, ui: &mut egui::Ui) 
    {
        egui::Frame::group(ui.style()).show(ui, |ui| 
        {
            ui.label(common::REF_FILE_HELP);
            ui.add_space(4.0);
            ui.horizontal(|ui| {
                ui.label(common::LABEL_FILE);
                ui.text_edit_singleline(&mut self.reference_section.path);
                if ui.button(common::BUTTON_BROWSE).clicked() {
                    if let Some(path_buf) = FileDialog::new().pick_file() {
                        if let Some(path_str) = path_buf.to_str() {
                            self.reference_section.path = path_str.to_string();
                            self.get_sheets_list(path_str)
                                .map(|sheets| self.reference_section.reference_sheet = sheets)
                                .map_err(|err| self.error = err)
                                .ok();
                        }
                    }
                }
            });

            ui.add_space(8.0);
            ui.label(common::REF_SHEET_HELP);
            ui.text_edit_singleline(&mut self.reference_section.reference_sheet);

            ui.add_space(4.0);
            ui.label(common::REF_KEY_COL_HELP);
            ui.text_edit_singleline(&mut self.reference_section.col_key);

            ui.add_space(4.0);
            ui.label(common::REF_VALUE_COL_HELP);
            ui.text_edit_singleline(&mut self.reference_section.col_value);

            ui.add_space(47.0);
            if ui.button(common::BUTTON_RUN_UPDATES).clicked()
            {
                self.error.clear();
                self.output_text.clear();

                let ref_sheets: Vec<String> = self.reference_section.reference_sheet.split(',').map(str::trim).map(String::from).collect();
                
                if 1 == ref_sheets.len() 
                {
                    // cargo run --bin rexcell -- -c cmd-update-sheets -t ../Test_Excell_new.xlsx -s C -d B -u "Лист1,Лист2,Лист3" -r ../Test_Excell_new.xlsx -e "Test" -k B -v C -i
                    let cfg: common::Config = common::Config {
                        command: common::Command::CmdUpdateSheets,
                        tgt_file: self.target_section.path.clone(), 
                        tgt_upd_table: self.target_section.update_sheet.clone(),
                        tgt_src_col: self.target_section.src_col.clone(),
                        tgt_dest_col: self.target_section.dest_col.clone(),
                        ref_file: self.reference_section.path.clone(),
                        ref_table: self.reference_section.reference_sheet.clone(),
                        ref_col_key: self.reference_section.col_key.clone(),
                        ref_col_value: self.reference_section.col_value.clone(),
                        new_sheet_name: self.target_section.new_file_name.clone(),
                        inplace: false,
                    };

                    let res = rexcell::execute(&cfg);

                    match res {
                        Ok(lines) => {
                            for line in &lines.0 {
                                self.output_text.push_str(line);
                                self.output_text.push_str("\n");
                            }
                            for line in &lines.1 {
                                self.output_text.push_str(line);
                                self.output_text.push_str("\n");
                            }
                            if cfg.inplace {
                                self.output_text.push_str(format!("Updated {} lines. {}\n", lines.0.len(), common::formatted_done_saved(&cfg.tgt_file)).as_str());
                            } else {
                                let new_file = format!("{}{}", cfg.tgt_file.trim_end_matches(common::XLSX_EXTENSION), common::NEW_FILE_SUFFIX);
                                self.output_text.push_str(format!("Updated {} lines. {}\n", lines.0.len(), common::formatted_done_saved(&new_file)).as_str());
                            }
                        }
                        Err(err) => {
                            self.error = format!("Failed to update {}: {}", cfg.tgt_file, err);
                        }
                    }
                }
                else
                {
                    self.error = String::from(common::ERROR_MULTIPLE_REF_SHEETS);
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
                                    self.draw_target_section(&mut columns[0]);
                                });
                            });
                    }
                    
                    Tab::Update => 
                    {
                        egui::Frame::group(ui.style()).show(ui, |ui| 
                        {
                            ui.columns(2, |columns| 
                            {
                                self.draw_target_section(&mut columns[0]);
                                self.draw_reference_section(&mut columns[1]);
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