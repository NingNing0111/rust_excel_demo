use read_bin::{get_input_data_list, get_rate_center_data};
use umya_spreadsheet::Spreadsheet;
use uuid::Uuid;

fn out_col(raws: Vec<String>, book: &mut Spreadsheet) {
    let uuid = Uuid::new_v4().to_string();
    let path = format!("{}.xlsx", uuid);
    let path = std::path::Path::new(&path);

    for i in 0..raws.len() {
        book.get_sheet_by_name_mut("Sheet1")
            .unwrap()
            .get_cell_mut((1 as u32, (i + 1) as u32))
            .set_value(raws[i].clone().as_str());
    }

    let _ = umya_spreadsheet::writer::xlsx::write(&book, path);
}

/// 根据公司名 打印利润中心编码
fn main() {
    let data = get_input_data_list();
    let raw_data = get_rate_center_data();
    let mut handle_data = Vec::<String>::new();
    for c_code in data {
        for raw in raw_data.iter() {
            if c_code == raw.company_code {
                handle_data.push(raw.rc_code.clone());
                break;
            }
        }
    }
    let mut book = umya_spreadsheet::new_file();
    out_col(handle_data, &mut book);
}
