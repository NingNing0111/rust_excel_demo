use std::{fs::File, io::Write, path::Path};

use read_bin::get_input_data_list;

fn main() {
    // 生成语句
    let keys = get_input_data_list();
    let values = get_input_data_list();
    let mut file = File::create(Path::new("keys_rs.txt")).unwrap();
    for i in 0..keys.len() {
        let s = format!("\"{}\".to_string(),\n", keys[i]);
        file.write(s.as_bytes()).unwrap();
    }

    let mut file = File::create(Path::new("values_rs.txt")).unwrap();
    for i in 0..values.len() {
        let s = format!("\"{}\".to_string(),\n", values[i]);
        file.write(s.as_bytes()).unwrap();
    }
}
