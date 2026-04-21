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
    #[arg(short = 's', long = "src_col", default_value = "C")]
    src_col: String,

    /// Колона за запис на уникалните номера
    #[arg(short = 'd', long = "dest_col", default_value = "B")]
    dest_col: String,

    /// Презаписва входния файл вместо създаване на нов
    #[arg(short = 'i', long = "inplace")]
    inplace: bool,
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

    let column_index = column_to_index(&args.src_col);
    let output_column_index = column_to_index(&args.dest_col);

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

    // 5. Запазване на промените
    if args.inplace {
        let _ = writer::xlsx::write(&book, path).expect("Грешка при запис");
        println!("Готово! Резултатът е записан в '{}'", args.file);
    } else {
        let new_file = format!("{}_new.xlsx", args.file.trim_end_matches(".xlsx"));
        let _ = writer::xlsx::write(&book, std::path::Path::new(&new_file)).expect("Грешка при запис");
        println!("Готово! Резултатът е записан в '{}'", new_file);
    }
}