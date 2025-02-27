use read_bin::{get_input_data_list, out_double_row};

/// 两行打印
fn main() {
    let data = get_input_data_list();
    println!("数据:{:?}", data);
    let mut book = umya_spreadsheet::new_file();
    out_double_row(data, &mut book);
}
