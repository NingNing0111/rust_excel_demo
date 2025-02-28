use read_bin::{find_rate_raw_by_c_name, get_input_data_list, out_double_row, Raw};

/// 根据公司名 得到公司代码 并且两行打印
fn main() {
    let c_names = get_input_data_list();

    let cleaned_names: Vec<String> = c_names
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

    let mut unknown_names: Vec<String> = Vec::new();

    let data: Vec<String> = cleaned_names
        .into_iter()
        .map(|name| {
            let raw = match find_rate_raw_by_c_name(name.clone()) {
                Some(raw) => raw,
                None => {
                    unknown_names.push(name.clone());
                    Raw {
                        rc_code: String::new(),
                        rc_name: String::new(),
                        company_code: String::new(),
                        company_name: String::new(),
                    }
                }
            };
            raw.company_code
        })
        .collect();

    println!("all data:{:?}", data);
    println!("all unknown names:{:?}", unknown_names);
    let mut book = umya_spreadsheet::new_file();
    out_double_row(data, &mut book);
}
