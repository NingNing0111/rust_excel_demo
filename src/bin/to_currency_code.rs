use read_bin::{find_rate_raw_by_c_name, get_input_data_list, Raw};
use uuid::Uuid;

fn main() {
    let c_names = get_input_data_list();
    let currency_codes = get_input_data_list();
    let cleaned_names: Vec<String> = c_names
        .clone()
        .into_iter()
        .map(|name| {
            let s = name.clone();
            let s = s.replace(" ", "");
            if let Some(index) = s.rfind('（') {
                // 找到最后一个 '（' 的位置
                s[..index].to_string() // 截取到 '(' 之前的部分
            } else if let Some(index) = s.rfind('(') {
                s[..index].to_string()
            } else {
                s // 如果没有 '（'，直接返回原字符串
            }
        })
        .collect();
    let c_codes: Vec<String> = cleaned_names
        .into_iter()
        .map(|name| {
            let raw = match find_rate_raw_by_c_name(name.clone()) {
                Some(raw) => raw,
                None => Raw {
                    rc_code: String::new(),
                    rc_name: String::new(),
                    company_code: String::new(),
                    company_name: String::new(),
                },
            };
            raw.company_code
        })
        .collect();

    let mut file = umya_spreadsheet::new_file();
    let book = &mut file;

    let path = format!("{}.xlsx", Uuid::new_v4().to_string());

    if c_names.len() == currency_codes.len() {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        for i in 0..c_names.len() {
            sheet
                .get_cell_mut((1 as u32, (2 * i + 1) as u32))
                .set_value(c_codes[i].clone().as_str());

            sheet
                .get_cell_mut((1 as u32, (2 * i + 2) as u32))
                .set_value(c_codes[i].clone().as_str());
            sheet
                .get_cell_mut((2 as u32, (2 * i + 1) as u32))
                .set_value(c_names[i].clone().as_str());

            sheet
                .get_cell_mut((2 as u32, (2 * i + 2) as u32))
                .set_value(c_names[i].clone().as_str());

            sheet
                .get_cell_mut((3 as u32, (2 * i + 1) as u32))
                .set_value(currency_codes[i].clone().as_str());

            sheet
                .get_cell_mut((3 as u32, (2 * i + 2) as u32))
                .set_value(currency_codes[i].clone().as_str());

            if c_names[i].contains("跨境人民币") && currency_codes[i] == String::from("CNY") {
                sheet
                    .get_cell_mut((4 as u32, (2 * i + 1) as u32))
                    .set_value("true");

                sheet
                    .get_cell_mut((4 as u32, (2 * i + 2) as u32))
                    .set_value("true");
            }
        }

        let _ = umya_spreadsheet::writer::xlsx::write(&book, path);
    } else {
        panic!("read faild!");
    }

    for i in 0..c_names.len() {
        println!("{} - {} - {}", c_names[i], currency_codes[i], c_codes[i]);
    }
}
