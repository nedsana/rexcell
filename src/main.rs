use umya_spreadsheet::*;
use clap::Parser;

#[derive(Parser)]
#[command(name = "rexcell")]
#[command(about = "Обработка на Excel файл с уникални ID-та")]
struct Args {
    /// Път до Excel файла
    #[arg(short, long, default_value = "data.xlsx")]
    file: String,

    /// Колона за търсене на повтарящи се текстове
    #[arg(short = 't', long, default_value = "C")]
    target_cell: String,

    /// Колона за запис на уникалните номера
    #[arg(short = 'i', long, default_value = "B")]
    id_cell: String,
}

fn column_to_index(col: &str) -> u32 {
    let mut index = 0;
    for c in col.chars() {
        index = index * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    index
}

fn main() {
    let args = Args::parse();

    let column_index = column_to_index(&args.target_cell);
    let output_column_index = column_to_index(&args.id_cell);

    // 1. Зареждане на съществуващ файл
    let path = std::path::Path::new(&args.file);
    let mut book = reader::xlsx::read(path).expect("Грешка при четене на файла");

    // 2. Избор на първия работен лист
    let sheet = book.get_sheet_mut(&0).expect("Листът не е намерен");

    // Използваме HashMap за проследяване на уникалните текстове и техните номера
    use std::collections::HashMap;
    let mut seen_texts: HashMap<String, i32> = HashMap::new();
    let mut counter = 1;

    // 3. Обхождане на редовете (от 1 до последния използван ред)
    let max_row = sheet.get_highest_row();
    
    for row in 1..=max_row {
        // Вземаме стойността от зададената колона
        let cell_value = sheet.get_value((column_index, row));
        
        if !cell_value.is_empty() {
            // Проверяваме дали вече сме виждали този текст
            let id = *seen_texts.entry(cell_value.clone()).or_insert_with(|| {
                let current_id = counter;
                counter += 1;
                current_id
            });

            // 4. Записваме уникалния номер в предишната колона
            sheet.get_cell_mut((output_column_index, row)).set_value(id.to_string());
            
            println!("Ред {}: Текст '{}' -> ID {}", row, cell_value, id);
        }
    }

    // 5. Запазване на промените (може и в нов файл за безопасност)
    let _ = writer::xlsx::write(&book, std::path::Path::new("result.xlsx")).expect("Грешка при запис");
    println!("Готово! Резултатът е записан в 'result.xlsx'");
}