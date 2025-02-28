use read_bin::{find_rate_raw_by_c_name, get_input_data_list, out_one_row, Raw};

fn main() {
    let c_names = get_input_data_list();

    let cleaned_names: Vec<String> = c_names
        .into_iter()
        .map(|name| {
            if let Some(index) = name.rfind('（') {
                // 找到最后一个 '（' 的位置
                name[..index].to_string() // 截取到 '(' 之前的部分
            } else if let Some(index) = name.rfind('(') {
                name[..index].to_string()
            } else {
                name // 如果没有 '（'，直接返回原字符串
            }
        })
        .collect();

    let mut unknown_names: Vec<String> = Vec::new();

    let data: Vec<String> = cleaned_names
        .clone()
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

    let mut book = umya_spreadsheet::new_file();
    out_one_row(cleaned_names, data, &mut book);
}
