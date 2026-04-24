use eframe::{egui, NativeOptions};
use rexcell::get_worksheet_names_string;
use umya_spreadsheet::reader;

struct GuiApp {
    path: String,
    sheet_names: String,
    error: String,
}

impl Default for GuiApp {
    fn default() -> Self {
        Self {
            path: String::from("data.xlsx"),
            sheet_names: String::new(),
            error: String::new(),
        }
    }
}

impl eframe::App for GuiApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("rexcell GUI");
            ui.label("Enter the Excel file path and click Load.");

            ui.horizontal(|ui| {
                ui.label("File:");
                ui.text_edit_singleline(&mut self.path);
                if ui.button("Load").clicked() {
                    self.error.clear();
                    self.sheet_names.clear();

                    let path = std::path::Path::new(&self.path);
                    match reader::xlsx::read(path) {
                        Ok(book) => {
                            self.sheet_names = get_worksheet_names_string(&book);
                        }
                        Err(err) => {
                            self.error = format!("Failed to load workbook: {}", err);
                        }
                    }
                }
            });

            if !self.error.is_empty() {
                ui.colored_label(egui::Color32::RED, &self.error);
            }

            if !self.sheet_names.is_empty() {
                ui.separator();
                ui.label("Worksheets:");
                ui.monospace(&self.sheet_names);
            }
        });
    }
}

fn main() {
    let options = NativeOptions::default();
    eframe::run_native("rexcell GUI", options, Box::new(|_cc| Box::new(GuiApp::default()))).expect("Failed to start GUI");
}
