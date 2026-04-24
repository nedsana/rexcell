use eframe::{egui, NativeOptions};
use rexcell::get_worksheet_names_string;
use rfd::FileDialog;
use umya_spreadsheet::reader;

struct GuiApp {
    path: String,
    text_a: String,
    text_b: String,
    text_c: String,
    output_text: String,
    error: String,
}

impl Default for GuiApp {
    fn default() -> Self {
        Self {
            path: String::from("data.xlsx"),
            text_a: String::new(),
            text_b: String::new(),
            text_c: String::new(),
            output_text: String::new(),
            error: String::new(),
        }
    }
}

impl GuiApp {
    fn draw_column(&mut self, ui: &mut egui::Ui) {
        egui::Frame::group(ui.style()).show(ui, |ui| {
            ui.label("File browser");
            ui.add_space(4.0);
            ui.horizontal(|ui| {
                ui.label("File:");
                ui.text_edit_singleline(&mut self.path);
                if ui.button("Browse").clicked() {
                    if let Some(path_buf) = FileDialog::new().pick_file() {
                        if let Some(path_str) = path_buf.to_str() {
                            self.path = path_str.to_string();
                        }
                    }
                }
            });

            ui.add_space(8.0);
            ui.label("Text field 1");
            ui.text_edit_singleline(&mut self.text_a);

            ui.add_space(4.0);
            ui.label("Text field 2");
            ui.text_edit_singleline(&mut self.text_b);

            ui.add_space(4.0);
            ui.label("Text field 3");
            ui.text_edit_singleline(&mut self.text_c);
        });
    }
}

impl eframe::App for GuiApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default().show(ctx, |ui| {
            ui.vertical(|ui| {
                ui.heading("rexcell GUI");
                ui.label("The top section has two identical panels.");
                ui.add_space(8.0);

                egui::Frame::group(ui.style()).show(ui, |ui| {
                    ui.columns(2, |columns| {
                        self.draw_column(&mut columns[0]);
                        self.draw_column(&mut columns[1]);
                    });
                });

                ui.add_space(12.0);

                egui::Frame::group(ui.style()).show(ui, |ui| {
                    ui.label("Output text");
                    ui.add_space(4.0);
                    ui.add(
                        egui::TextEdit::multiline(&mut self.output_text)
                            .desired_rows(12)
                            .desired_width(f32::INFINITY)
                            .lock_focus(true),
                    );
                });

                if ui.button("Load workbook").clicked() {
                    self.error.clear();
                    self.output_text.clear();

                    let path = std::path::Path::new(&self.path);
                    match reader::xlsx::read(path) {
                        Ok(book) => {
                            self.output_text = get_worksheet_names_string(&book);
                        }
                        Err(err) => {
                            self.error = format!("Failed to load workbook: {}", err);
                        }
                    }
                }

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
    eframe::run_native("rexcell GUI", options, Box::new(|_cc| Box::new(GuiApp::default()))).expect("Failed to start GUI");
}
