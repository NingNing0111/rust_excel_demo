use std::{fs::File, io::Write};

use read_bin::{constant::get_unit, get_input_data_list, out_one_row};

fn main() {
    // 科目编号
    let kmbh = get_input_data_list();
    // 往来单位编号
    let data = get_input_data_list();
    let mut w_data = Vec::<String>::new();
    let mut e_data = Vec::<usize>::new();
    for i in 0..data.len() {
        if !data[i].is_empty() {
            let res = get_unit(data[i].clone());
            w_data.push(res);
        } else {
            e_data.push(i);
            w_data.push(String::new());
        }
    }

    let mut book = umya_spreadsheet::new_file();
    out_one_row(data, w_data, &mut book);

    // 为空的记录
    if !e_data.is_empty() {
        let mut file = File::create("empty_jy.txt").unwrap();

        let mut e_code = Vec::<String>::new();
        for i in e_data {
            if !e_code.contains(&kmbh[i]) {
                e_code.push(kmbh[i].clone());
            }
        }

        file.write("原FMIS ".as_bytes()).unwrap();
        for i in e_code {
            let s = format!("{}、", i);
            file.write(s.as_bytes()).unwrap();
        }
        file.write("往来单位核算维度为空".as_bytes()).unwrap();
    }
}
