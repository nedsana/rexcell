use eframe::{egui, NativeOptions};
use rfd::FileDialog;
// use std::process::Command;
use rexcell::common;
use rexcell::excell;
use std::sync::mpsc;
use std::thread;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};

use log::{debug, info, error};

#[derive(Debug, Clone)]
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

#[derive(Debug, Clone)]
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
            col_key: String::from(common::REF_DEFAULT_SRC_COL),
            col_value: String::from(common::REF_DEFAULT_DST_COL),
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

#[derive(Debug, Clone)]
struct AnalysisData { 
    analysis_path: String,
    analysis_sheet: String,
}

impl Default for AnalysisData {
    fn default() -> Self {
        Self {
            analysis_path: String::from(common::REF_DEFAULT_EXCEL_FILE),
            analysis_sheet: String::from(common::REF_DEFAULT_TABLE),
        }
    }
}

impl AnalysisData {
    pub fn new(p_path: String, p_reference_sheet: String) -> Self {
        Self { 
            analysis_path: String::from(p_path),
            analysis_sheet: String::from(p_reference_sheet),
        }
    }
}

#[derive(PartialEq)]
enum Tab 
{
    Filter,
    Update,
    AutoComplete
}

struct GuiApp 
{
    cfg_filter: TargetData,

    cfg_update_tgt: TargetData,
    cfg_update_ref: ReferencesData,

    analysis_data: AnalysisData,

    error: String,

    active_tab: Tab,

    log_buffer: String, // Buffer to hold the logs for the GUI field
    is_working: Arc<AtomicBool>,   // Flag to indicate if processing is on
    log_rx: mpsc::Receiver<String>, // Channel to receive logs
}

impl Default for GuiApp 
{
    fn default() -> Self 
    {
        let (_, log_rx_tmp) = mpsc::channel::<String>();

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
                                    common::REF_DEFAULT_SRC_COL.to_string(),
                                    common::REF_DEFAULT_DST_COL.to_string()),

            analysis_data: AnalysisData::new( common::REF_DEFAULT_EXCEL_FILE.to_string(), 
                                    common::REF_DEFAULT_TABLE.to_string()),

            error: String::new(),

            active_tab: Tab::Filter,

            log_buffer: String::default(),

            is_working: Arc::new(AtomicBool::new(false)),
          
            log_rx: log_rx_tmp,
        }
    }
}

impl GuiApp 
{
    fn new(_cc: &eframe::CreationContext<'_>) -> Self 
    {
        //Create the actual channel to connect to fern
        let (log_tx, log_rx) = mpsc::channel::<String>();
        // let ctx_clone = _cc.egui_ctx.clone();

        //Init the fern logger
        fern::Dispatch::new()
            .format(|out, message, record| 
            {
                let level_str = format!("{:<5}", record.level().to_string());

                let file_line = format!("{}:{}", record.file().unwrap_or("unknown"), record.line().unwrap_or(0));

                let file_line_padded = format!("{:<32}", file_line);

                out.finish(format_args!("[{}] [{}] {}", level_str, file_line_padded, message))
            })
            .level(log::LevelFilter::Debug)
            // .chain(std::io::stdout())
            .chain(fern::Output::call(move |record| 
            {
                let _ = log_tx.send(format!("{}\n", record.args()));
                // ctx_clone.request_repaint(); // Should Wake-up GUI, but actually blocks GUI
            }))
            .apply()
            .unwrap();

        info!("Logging setup complete!");

        //Use the default constructor for all fields, just update log_rx
        Self {
            log_rx,
            ..Self::default()
        }
    }

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

    fn draw_filter_section(&mut self, ui: &mut egui::Ui, headers: &[&str], filtering: bool)
    {
        egui::Frame::group(ui.style()).show(ui, |ui| 
        {
            let tgt_data: &mut TargetData = if filtering {
                &mut self.cfg_filter
            } else {
                &mut self.cfg_update_tgt
            };

            ui.label(headers[0]);
            ui.add_space(4.0);
            ui.horizontal(|ui| {
                ui.label(headers[1]);
                ui.text_edit_singleline(&mut tgt_data.path);
                if ui.button(headers[2]).clicked() {
                    if let Some(path_buf) = FileDialog::new().pick_file() {
                        if let Some(path_str) = path_buf.to_str() {
                            tgt_data.path = path_str.to_string();
                            Self::get_sheets_list(path_str)
                                .map(|sheets| tgt_data.update_sheets = sheets)
                                .map_err(|err| self.error = err)
                                .ok();
                        }
                    }
                }
            });

            ui.add_space(8.0);
            ui.label(headers[3]);
            ui.text_edit_singleline(&mut tgt_data.update_sheets);

            ui.add_space(4.0);
            ui.label(headers[4]);
            ui.text_edit_singleline(&mut tgt_data.src_col);

            ui.add_space(4.0);
            ui.label(headers[5]);
            ui.text_edit_singleline(&mut tgt_data.dest_col);
            
            ui.add_space(4.0);
            if headers.len() > 6
            {
                ui.add_space(4.0);
                ui.label(headers[6]);
                ui.text_edit_singleline(&mut tgt_data.new_sheet_name);

                if ui.button(headers[7]).clicked()
                {
                    if false == self.is_working.load(Ordering::SeqCst)
                    {
                        // cargo run --bin rexcell -- -c cmd-filter-sheets -t ../Test_Excell.xlsx -u "Лист1,Лист2,Лист3" -s C -d E -n "Test"
                        let cfg: common::Config = common::Config {
                            command: common::Command::CmdFilterSheets,
                            tgt_file: tgt_data.path.clone(), 
                            tgt_upd_table: tgt_data.update_sheets.clone(),
                            tgt_src_col: tgt_data.src_col.clone(),
                            tgt_dest_col: tgt_data.dest_col.clone(),
                            ref_file: "".to_string(),
                            ref_table: "".to_string(),
                            ref_col_key: "".to_string(),
                            ref_col_value: "".to_string(),
                            new_sheet_name: tgt_data.new_sheet_name.clone(),
                            inplace: true,
                            analysis_file: "".to_string(),
                            analysis_table: "".to_string(),
                        };
                        
                        debug!("Start filtering!");

                        self.is_working.store(true, Ordering::SeqCst);
                        let is_working_clone = self.is_working.clone();

                        thread::spawn(move || 
                        {
                            let res = excell::execute(&cfg);
                            match res 
                            {
                                Ok(()) => 
                                {
                                    info!("Filtering file {} - OK!", cfg.tgt_file);
                                }
                                Err(err) => 
                                {
                                    error!("Filtering file {} - FAILED! {}", cfg.tgt_file, err);
                                    // self.error = format!("Failed to update file {}! {}\n", cfg.tgt_file, out.1);
                                }
                            }
                            is_working_clone.store(false, Ordering::SeqCst);
                        });
                    }
                    else 
                    {
                        debug!("Filtering is running!");
                    }
                }
            }
        });
    }

    fn draw_cfg_update_ref(&mut self, ui: &mut egui::Ui, headers: &[&str]) 
    {
        egui::Frame::group(ui.style()).show(ui, |ui| 
        {
            ui.label(headers[0]);
            ui.add_space(4.0);
            ui.horizontal(|ui| {
                ui.label(headers[1]);
                ui.text_edit_singleline(&mut self.cfg_update_ref.path);
                if ui.button(headers[2]).clicked() {
                    if let Some(path_buf) = FileDialog::new().pick_file() {
                        if let Some(path_str) = path_buf.to_str() {
                            self.cfg_update_ref.path = path_str.to_string();
                            Self::get_sheets_list(path_str)
                                .map(|sheets| self.cfg_update_ref.reference_sheet = sheets)
                                .map_err(|err| self.error = err)
                                .ok();
                        }
                    }
                }
            });

            ui.add_space(8.0);
            ui.label(headers[3]);
            ui.text_edit_singleline(&mut self.cfg_update_ref.reference_sheet);

            ui.add_space(4.0);
            ui.label(headers[4]);
            ui.text_edit_singleline(&mut self.cfg_update_ref.col_key);

            ui.add_space(4.0);
            ui.label(headers[5]);
            ui.text_edit_singleline(&mut self.cfg_update_ref.col_value);

            ui.add_space(4.0);
            if ui.button(headers[6]).clicked()
            {
                if false == self.is_working.load(Ordering::SeqCst)
                {
                    let ref_sheets: Vec<String> = self.cfg_update_ref.reference_sheet.split(',').map(str::trim).map(String::from).collect();
                    
                    if 1 == ref_sheets.len() 
                    {
                        // cargo run --bin rexcell -- -c cmd-update-sheets -t ../Test_Excell_new.xlsx -s C -d B -u "Лист1,Лист2,Лист3" -r ../Test_Excell_new.xlsx -e "Test" -k B -v C -i
                        let cfg: common::Config = common::Config {
                            command: common::Command::CmdUpdateSheets,
                            tgt_file: self.cfg_update_tgt.path.clone(), 
                            tgt_upd_table: self.cfg_update_tgt.update_sheets.clone(),
                            tgt_src_col: self.cfg_update_tgt.src_col.clone(),
                            tgt_dest_col: self.cfg_update_tgt.dest_col.clone(),
                            ref_file: self.cfg_update_ref.path.clone(),
                            ref_table: self.cfg_update_ref.reference_sheet.clone(),
                            ref_col_key: self.cfg_update_ref.col_key.clone(),
                            ref_col_value: self.cfg_update_ref.col_value.clone(),
                            new_sheet_name: self.cfg_update_tgt.new_sheet_name.clone(),
                            inplace: true,
                            analysis_file: "".to_string(),
                            analysis_table: "".to_string(),
                        };

                        debug!("Start updating!");

                        self.is_working.store(true, Ordering::SeqCst);
                        let is_working_clone = self.is_working.clone();

                        thread::spawn(move || 
                        {
                            let res = excell::execute(&cfg);
                            match res 
                            {
                                Ok(()) => 
                                {
                                    info!("Updating file {} - OK!", cfg.tgt_file);
                                }
                                Err(err) => 
                                {
                                    error!("Updating file {} - FAILED! {}", cfg.tgt_file, err);
                                    // self.error = format!("Failed to update file {}! {}\n", cfg.tgt_file, out.1);
                                }
                            }
                            is_working_clone.store(false, Ordering::SeqCst);
                        });
                    }
                    else
                    {
                        self.error = String::from(common::ERROR_MULTIPLE_REF_SHEETS);
                    }
                }
                else
                {
                    info!("Updating is running!");
                }
            }
        });
    }

    fn draw_autocomplete_section(&mut self, ui: &mut egui::Ui, headers: &[&str], filtering: bool)
    {
        egui::Frame::group(ui.style()).show(ui, |ui| 
        {
            let tgt_data: &mut TargetData = if filtering {
                &mut self.cfg_filter
            } else {
                &mut self.cfg_update_tgt
            };

            ui.label(headers[0]);
            ui.add_space(4.0);
            ui.horizontal(|ui| {
                ui.label(headers[1]);
                ui.text_edit_singleline(&mut tgt_data.path);
                if ui.button(headers[2]).clicked() {
                    if let Some(path_buf) = FileDialog::new().pick_file() {
                        if let Some(path_str) = path_buf.to_str() {
                            tgt_data.path = path_str.to_string();
                            Self::get_sheets_list(path_str)
                                .map(|sheets| tgt_data.update_sheets = sheets)
                                .map_err(|err| self.error = err)
                                .ok();
                        }
                    }
                }
            });

            ui.add_space(8.0);
            ui.label(headers[3]);
            ui.text_edit_singleline(&mut tgt_data.update_sheets);

            ui.add_space(4.0);
            ui.label(headers[4]);
            ui.text_edit_singleline(&mut tgt_data.src_col);

            ui.add_space(4.0);
            ui.label(headers[5]);
            ui.text_edit_singleline(&mut tgt_data.dest_col);

            ui.add_space(4.0);
            ui.label(headers[6]);
            ui.text_edit_singleline(&mut tgt_data.new_sheet_name);

            // analysis_data
            ui.add_space(4.0);
            ui.label(headers[7]);
            ui.add_space(4.0);
            ui.horizontal(|ui| {
                ui.label(headers[8]);
                ui.text_edit_singleline(&mut self.analysis_data.analysis_path);
                if ui.button(headers[9]).clicked() {
                    if let Some(path_buf) = FileDialog::new().pick_file() {
                        if let Some(path_str) = path_buf.to_str() {
                            self.analysis_data.analysis_path = path_str.to_string();
                            Self::get_sheets_list(path_str)
                                .map(|sheets| self.analysis_data.analysis_sheet = sheets)
                                .map_err(|err| self.error = err)
                                .ok();
                        }
                    }
                }
            });

            ui.add_space(8.0);
            ui.label(headers[9]);
            ui.text_edit_singleline(&mut self.analysis_data.analysis_sheet);

            ui.add_space(4.0);
            if ui.button(headers[10]).clicked()
            {
                if false == self.is_working.load(Ordering::SeqCst)
                {
                    // cargo run --bin rexcell -- -c cmd-filter-sheets -t ../Test_Excell.xlsx -u "Лист1,Лист2,Лист3" -s C -d E -n "Test"
                    let cfg: common::Config = common::Config {
                        command: common::Command::CmdAutocompleteSheets,
                        tgt_file: tgt_data.path.clone(), 
                        tgt_upd_table: tgt_data.update_sheets.clone(),
                        tgt_src_col: tgt_data.src_col.clone(),
                        tgt_dest_col: tgt_data.dest_col.clone(),
                        ref_file: "".to_string(),
                        ref_table: "".to_string(),
                        ref_col_key: "".to_string(),
                        ref_col_value: "".to_string(),
                        new_sheet_name: tgt_data.new_sheet_name.clone(),
                        inplace: true,
                        analysis_file: self.analysis_data.analysis_path.clone(),
                        analysis_table: self.analysis_data.analysis_sheet.clone()
                    };
                    
                    debug!("Start autocomplete!");

                    self.is_working.store(true, Ordering::SeqCst);
                    let is_working_clone = self.is_working.clone();

                    thread::spawn(move || 
                    {
                        let res = excell::execute(&cfg);
                        match res 
                        {
                            Ok(()) => 
                            {
                                info!("Autocompleting file {} - OK!", cfg.tgt_file);
                            }
                            Err(err) => 
                            {
                                error!("Autocompleting file {} - FAILED! {}", cfg.tgt_file, err);
                                // self.error = format!("Failed to update file {}! {}\n", cfg.tgt_file, out.1);
                            }
                        }
                        is_working_clone.store(false, Ordering::SeqCst);
                    });
                }
                else 
                {
                    debug!("Autocomplete is running!");
                }
            }
        });
    }
}

const FILTER_SECTION_HEADERS: [&str; 8] = [common::TGT_FILE_HELP, common::LABEL_FILE, common::BUTTON_BROWSE, 
                                        common::LIST_SHEETS_TO_UPDATE, common::TGT_SRC_COL_HELP, common::TGT_DEST_COL_ACCUM_HELP, 
                                        common::NEW_SHEET_NAME_HELP, common::BUTTON_FILTER_DATA];

const AUTOCOMPLETE_SECTION_HEADERS: [&str; 11] = [common::TGT_FILE_HELP, common::LABEL_FILE, common::BUTTON_BROWSE, 
                                                 common::LIST_SHEETS_TO_UPDATE, common::TGT_SRC_COL_HELP, common::TGT_DEST_COL_ACCUM_HELP, common::NEW_SHEET_NAME_HELP, 
                                                 common::ANALYSIS_FILE_HELP, common::LABEL_FILE, common::BUTTON_BROWSE,
                                                 common::BUTTON_FILTER_DATA];

const UPDATE_SECTION_TGT_HEADERS: [&str; 6] = [common::TGT_FILE_HELP, common::LABEL_FILE, common::BUTTON_BROWSE, 
                                               common::LIST_SHEETS_TO_UPDATE, common::REF_SRC_COL_HELP, common::TGT_DEST_COL_HELP];

const UPDATE_SECTION_REF_HEADERS: [&str; 7] = [common::REF_FILE_HELP, common::LABEL_FILE, common::BUTTON_BROWSE, 
                                               common::REF_SHEET_HELP, common::REF_SRC_COL_HELP, common::REF_DEST_COL_HELP, 
                                               common::BUTTON_RUN_UPDATES];

impl eframe::App for GuiApp 
{
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) 
    {
        const MAX_HEIGHT: f32 = 400.0;
        const SCALE_FACTOR: f32 = 1.0;
        ctx.set_pixels_per_point(SCALE_FACTOR);

        //Get current logs and copy them to buffer
        while let Ok(new_log) = self.log_rx.try_recv() 
        {
            self.log_buffer.push_str(&new_log);
            
            // // Optional: limit the log buffer size, as this can delay the gui
            // if self.log_buffer.len() > 10_000 {
            //     self.log_buffer = self.log_buffer.chars().skip(2000).collect();
            // }
        }

        egui::CentralPanel::default().show(ctx, |ui| 
        {
            ui.vertical(|ui| 
            {
                ui.heading(common::WINDOW_TITLE);
                // ui.label(common::PANEL_DESCRIPTION);

                ui.add_space(8.0);

                ui.horizontal(|ui| 
                {
                    ui.selectable_value(&mut self.active_tab, Tab::AutoComplete, common::TAB_LABEL_AUTOCOMPLETE);
                    ui.selectable_value(&mut self.active_tab, Tab::Filter, common::TAB_LABEL_FILTER);
                    ui.selectable_value(&mut self.active_tab, Tab::Update, common::TAB_LABEL_UPDATE);
                });

                match self.active_tab 
                {
                    Tab::AutoComplete => 
                    {
                        egui::Frame::group(ui.style()).show(ui, |ui| 
                        {
                            ui.columns(2, |columns| 
                            {
                                self.draw_autocomplete_section(&mut columns[0], &AUTOCOMPLETE_SECTION_HEADERS, true);
                            });
                        });
                    }

                    Tab::Filter => 
                    {
                        egui::Frame::group(ui.style()).show(ui, |ui| 
                            {
                                ui.columns(2, |columns| 
                                {
                                    self.draw_filter_section(&mut columns[0], &FILTER_SECTION_HEADERS, true);
                                });
                            });
                    }
                    
                    Tab::Update => 
                    {
                        egui::Frame::group(ui.style()).show(ui, |ui| 
                        {
                            ui.columns(2, |columns| 
                            {
                                self.draw_filter_section(&mut columns[0], &UPDATE_SECTION_TGT_HEADERS, false);

                                self.draw_cfg_update_ref(&mut columns[1], &UPDATE_SECTION_REF_HEADERS);
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
                        .max_height(MAX_HEIGHT * SCALE_FACTOR) 
                        .auto_shrink([false; 2]) 
                        .stick_to_bottom(true)
                        .show(ui, |ui| {
                            ui.add(
                                egui::TextEdit::multiline(&mut self.log_buffer)
                                    .desired_rows(16)
                                    .desired_width(f32::INFINITY)
                                    .lock_focus(true)
                                    .interactive(true),
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
        Box::new(|cc| Box::new(GuiApp::new(cc)))).expect(common::ERROR_FAILED_TO_START_GUI);
}

// RUST_BACKTRACE=1 cargo run --bin gui >OUT
// RUST_BACKTRACE=1 cargo run --bin rexcell -- -t ../../Test_Twins.xlsx -e "Ед. Цени" -u "Ф200" -k B -v C -s C -d B -i