use std::path::Path;

use read_bin::{
    constant::get_unit, find_rate_raw_by_c_name, get_input_data_list, get_rate_center_data,
    read_line_from_file, Raw,
};
use umya_spreadsheet::Spreadsheet;

// 人民币和跨境人民币区分
fn yuan(book: &mut Spreadsheet, c_names: Vec<String>, currency_codes: Vec<String>) {
    let cleaned_names: Vec<String> = c_names
        .clone()
        .into_iter()
        .map(|name| {
            let mut s: String = name.clone();
            if name.contains(" ") {
                s = s.replace(" ", "");
            }
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
                    company_code: String::from(&format!("{}-未知公司代码", name)),
                    company_name: name,
                },
            };
            raw.company_code
        })
        .collect();

    // 利润中心
    let raw_data = get_rate_center_data();
    let mut rc_codes = Vec::<String>::new();
    for c_code in c_codes.clone() {
        let mut rc_code = String::new();
        let mut is_find = false;
        for raw in raw_data.iter() {
            if c_code == raw.company_code {
                rc_code = raw.rc_code.clone();
                is_find = true;
                break;
            }
        }
        if is_find {
            rc_codes.push(rc_code);
        } else {
            rc_codes.push(String::from(&format!("{} -> 未知利润中心", c_code)));
        }
    }

    if c_names.len() == currency_codes.len() {
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
        for i in 0..c_names.len() {
            // 公司编码
            sheet
                .get_cell_mut((1 as u32, (2 * i + 1) as u32))
                .set_value(c_codes[i].clone().as_str());

            sheet
                .get_cell_mut((1 as u32, (2 * i + 2) as u32))
                .set_value(c_codes[i].clone().as_str());

            // 公司名
            sheet
                .get_cell_mut((2 as u32, (2 * i + 1) as u32))
                .set_value(c_names[i].clone().as_str());

            sheet
                .get_cell_mut((2 as u32, (2 * i + 2) as u32))
                .set_value(c_names[i].clone().as_str());

            // 货币标识
            sheet
                .get_cell_mut((3 as u32, (2 * i + 1) as u32))
                .set_value(currency_codes[i].clone().as_str());

            sheet
                .get_cell_mut((3 as u32, (2 * i + 2) as u32))
                .set_value(currency_codes[i].clone().as_str());

            // 跨境人民币区分
            if c_names[i].contains("跨境人民币") && currency_codes[i] == String::from("CNY") {
                sheet
                    .get_cell_mut((4 as u32, (2 * i + 1) as u32))
                    .set_value("true");

                sheet
                    .get_cell_mut((4 as u32, (2 * i + 2) as u32))
                    .set_value("true");
            }

            // 利润中心
            sheet
                .get_cell_mut((5 as u32, (2 * i + 1) as u32))
                .set_value(rc_codes[i].clone().as_str());

            sheet
                .get_cell_mut((5 as u32, (2 * i + 2) as u32))
                .set_value(rc_codes[i].clone().as_str());
        }
    } else {
        panic!("read faild!");
    }
}

fn out_money(book: &mut Spreadsheet, money: Vec<String>) {
    let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
    for i in 0..money.len() {
        sheet
            .get_cell_mut((6 as u32, (2 * i + 1) as u32))
            .set_value(money[i].clone().as_str());
        sheet
            .get_cell_mut((6 as u32, (2 * i + 2) as u32))
            .set_value(money[i].clone().as_str());

        sheet
            .get_cell_mut((7 as u32, (2 * i + 1) as u32))
            .set_value(money[i].clone().as_str());
        sheet
            .get_cell_mut((7 as u32, (2 * i + 2) as u32))
            .set_value(money[i].clone().as_str());
        if money[i].starts_with("-") {
            sheet
                .get_cell_mut((8 as u32, (2 * i + 1) as u32))
                .set_value("50");
            sheet
                .get_cell_mut((8 as u32, (2 * i + 2) as u32))
                .set_value("40");
        } else {
            sheet
                .get_cell_mut((8 as u32, (2 * i + 1) as u32))
                .set_value("40");
            sheet
                .get_cell_mut((8 as u32, (2 * i + 2) as u32))
                .set_value("50");
        }
    }
}

// 交易单位
fn out_wl_gs(book: &mut Spreadsheet, wl_code: Vec<String>) {
    let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();
    for i in 0..wl_code.len() {
        sheet
            .get_cell_mut((10 as u32, (2 * i + 1) as u32))
            .set_value(wl_code[i].as_str());
        sheet
            .get_cell_mut((10 as u32, (2 * i + 2) as u32))
            .set_value(wl_code[i].as_str());
        if !wl_code[i].is_empty() {
            let res = get_unit(wl_code[i].clone());
            sheet
                .get_cell_mut((9 as u32, (2 * i + 1) as u32))
                .set_value(res.as_str());
            sheet
                .get_cell_mut((9 as u32, (2 * i + 2) as u32))
                .set_value(res.as_str());
        } else {
            sheet
                .get_cell_mut((9 as u32, (2 * i + 1) as u32))
                .set_value("");
            sheet
                .get_cell_mut((9 as u32, (2 * i + 2) as u32))
                .set_value("");
        }
    }
}

fn main() {
    let c_name_path = Path::new("base/原数据-责任中心名称.txt");
    let currency_code_path = Path::new("base/原数据-币种简称.txt");
    let money_path = Path::new("base/原数据-余额.txt");
    let wl_code_path = Path::new("base/原数据-往来单位编码.txt");
    let out_path = Path::new("base_result.xlsx");

    let c_names = read_line_from_file(c_name_path);
    let currency_code = read_line_from_file(currency_code_path);
    let money = read_line_from_file(money_path);
    let wl_codes = read_line_from_file(wl_code_path);

    let mut book = umya_spreadsheet::new_file();
    yuan(&mut book, c_names.clone(), currency_code.clone());
    out_money(&mut book, money.clone());
    if wl_codes.len() > 0 {
        out_wl_gs(&mut book, wl_codes.clone());
    }
    let _ = umya_spreadsheet::writer::xlsx::write(&book, out_path);

    println!("===============================================");
    println!(
        "c_name_len:{} 行 currency_len:{} 行 money:{} 行 wl_code_len:{} 行",
        c_names.len(),
        currency_code.len(),
        money.len(),
        wl_codes.len()
    );
    println!("===============================================");

    let _ = get_input_data_list();
}
