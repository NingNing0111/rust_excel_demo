use read_bin::{get_input_data_list, out_one_row};

fn main() {
    let data = get_input_data_list();
    let mut g_code = Vec::<String>::new();
    let len = data.len();
    let mut i = 0;
    while i < len / 2 {
        if data[2 * i].starts_with("-") {
            g_code.push("50".to_string());
            g_code.push("40".to_string());
        } else {
            g_code.push("40".to_string());
            g_code.push("50".to_string());
        }
        i += 1;
    }
    let mut book = umya_spreadsheet::new_file();
    out_one_row(data, g_code, &mut book);
}
