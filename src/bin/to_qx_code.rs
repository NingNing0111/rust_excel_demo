use std::path::Path;

use read_bin::{
    constant::{get_fixed, get_flow, get_wh_fixed, get_wh_flow},
    out_one_row, read_line_from_file,
};
use umya_spreadsheet::Spreadsheet;

fn out(book: &mut Spreadsheet, data: Vec<String>) {
    let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
    for i in 0..data.len() {
        sheet
            .get_cell_mut((1 as u32, (2 * i + 1) as u32))
            .set_value(data[i].as_str());
        sheet
            .get_cell_mut((1 as u32, (2 * i + 2) as u32))
            .set_value(data[i].as_str());
    }
}

fn main() {
    let ex_names_path = Path::new("base/原数据-期限名称.txt");
    let ex_codes_path = Path::new("base/原数据-期限编码.txt");
    let currency_codes_path = Path::new("base/原数据-币种简称.txt");
    let out_path = Path::new("期限结果.xlsx");

    let ex_names = read_line_from_file(ex_names_path);
    let ex_codes = read_line_from_file(ex_codes_path);
    let currency_codes = read_line_from_file(currency_codes_path);

    let mut book = umya_spreadsheet::new_file();
    let mut data = Vec::<String>::new();
    for i in 0..ex_names.len() {
        if ex_codes[i].is_empty() || ex_names[i].is_empty() {
            data.push(String::new());
            continue;
        }
        if currency_codes[i] != String::from("CNY") {
            // 流动
            if ex_codes[i].starts_with("411") {
                let flow_code = get_wh_flow(ex_names[i].clone().replace("个月", "").to_string());
                data.push(flow_code);
            } else {
                // 固定
                let fixed_code = get_wh_fixed(ex_names[i].clone().replace("个月", "").to_string());
                data.push(fixed_code);
            }
        } else {
            // 流动
            if ex_codes[i].starts_with("411") {
                let flow_code = get_flow(ex_names[i].clone().replace("个月", "").to_string());
                data.push(flow_code);
            } else {
                // 固定
                let fixed_code = get_fixed(ex_names[i].clone().replace("个月", "").to_string());
                data.push(fixed_code);
            }
        }
    }

    out(&mut book, data.clone());

    let _ = umya_spreadsheet::writer::xlsx::write(&book, out_path);

    let mut book = umya_spreadsheet::new_file();
    out_one_row(ex_names, data, &mut book);
}
