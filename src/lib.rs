use std::{
    fs::File,
    io::{self, Read, Write},
    path::Path,
};
pub mod constant;

use umya_spreadsheet::Spreadsheet;
use uuid::Uuid;

pub struct Raw {
    pub rc_code: String,
    pub rc_name: String,
    pub company_code: String,
    pub company_name: String,
}

pub fn get_rate_center_data() -> Vec<Raw> {
    let data = vec![
        Raw {
            rc_code: "H200700001".to_string(),
            rc_name: "中油财务有限责任公司北京总部".to_string(),
            company_code: "H200".to_string(),
            company_name: "中油财务有限责任公司".to_string(),
        },
        Raw {
            rc_code: "H201700001".to_string(),
            rc_name: "中油财务有限责任公司大庆分公司".to_string(),
            company_code: "H201".to_string(),
            company_name: "中油财务有限责任公司大庆分公司".to_string(),
        },
        Raw {
            rc_code: "H202700001".to_string(),
            rc_name: "中油财务有限责任公司沈阳分公司".to_string(),
            company_code: "H202".to_string(),
            company_name: "中油财务有限责任公司沈阳分公司".to_string(),
        },
        Raw {
            rc_code: "H203700001".to_string(),
            rc_name: "中油财务有限责任公司吉林分公司".to_string(),
            company_code: "H203".to_string(),
            company_name: "中油财务有限责任公司吉林分公司".to_string(),
        },
        Raw {
            rc_code: "H204700001".to_string(),
            rc_name: "中油财务有限责任公司西安分公司".to_string(),
            company_code: "H204".to_string(),
            company_name: "中油财务有限责任公司西安分公司".to_string(),
        },
        Raw {
            rc_code: "H205700001".to_string(),
            rc_name: "中国石油财务（香港）有限公司".to_string(),
            company_code: "H205".to_string(),
            company_name: "中国石油财务（香港）有限公司".to_string(),
        },
        Raw {
            rc_code: "H205700001".to_string(),
            rc_name: "中国石油财务(香港)有限公司".to_string(),
            company_code: "H205".to_string(),
            company_name: "中国石油财务(香港)有限公司".to_string(),
        },
        Raw {
            rc_code: "H206700001".to_string(),
            rc_name: "中国石油财务（迪拜）有限公司".to_string(),
            company_code: "H206".to_string(),
            company_name: "中国石油财务（迪拜）有限公司".to_string(),
        },
        Raw {
            rc_code: "H206700001".to_string(),
            rc_name: "中国石油财务(迪拜)有限公司".to_string(),
            company_code: "H206".to_string(),
            company_name: "中国石油财务(迪拜)有限公司".to_string(),
        },
        Raw {
            rc_code: "H207700001".to_string(),
            rc_name: "中国石油财务（新加坡）有限公司".to_string(),
            company_code: "H207".to_string(),
            company_name: "中国石油财务（新加坡）有限公司".to_string(),
        },
        Raw {
            rc_code: "H207700001".to_string(),
            rc_name: "中国石油财务(新加坡)有限公司".to_string(),
            company_code: "H207".to_string(),
            company_name: "中国石油财务(新加坡)有限公司".to_string(),
        },
        Raw {
            rc_code: "H208700001".to_string(),
            rc_name: "中油财务有限责任公司合并抵销中心".to_string(),
            company_code: "H208".to_string(),
            company_name: "CNPC (HK) Overseas Capital Ltd.".to_string(),
        },
        Raw {
            rc_code: "H209700001".to_string(),
            rc_name: "总部和分公司抵消".to_string(),
            company_code: "H209".to_string(),
            company_name: "CNPC Golden Autumn Limited".to_string(),
        },
        Raw {
            rc_code: "H210700001".to_string(),
            rc_name: "总部和子公司抵消".to_string(),
            company_code: "H210".to_string(),
            company_name: "CNPC (BVI) Limited.".to_string(),
        },
    ];
    data
}

pub fn find_rate_raw_by_c_name(name: String) -> Option<Raw> {
    let data = get_rate_center_data();
    for raw in data {
        if raw.rc_name == name {
            return Some(raw);
        }
    }
    None
}

pub fn get_input_data_list() -> Vec<String> {
    io::stdout().flush().unwrap();
    println!("请输入数据 换行符隔开 exit 结束:");
    let mut data = Vec::<String>::new();

    loop {
        let mut raw = String::new();

        io::stdout().flush().unwrap();
        io::stdin().read_line(&mut raw).unwrap();
        raw = raw.trim().to_string();
        println!("输入：{}", raw);

        if raw == String::from("exit") {
            break;
        } else {
            data.push(raw);
        }
    }

    data
}

pub fn out_double_row(raws: Vec<String>, book: &mut Spreadsheet) {
    let uuid = Uuid::new_v4().to_string();
    let path = format!("{}.xlsx", uuid);
    let path = std::path::Path::new(&path);

    for i in 0..raws.len() {
        book.get_sheet_by_name_mut("Sheet1")
            .unwrap()
            .get_cell_mut((1 as u32, (2 * i + 1) as u32))
            .set_value(raws[i].clone().as_str());

        book.get_sheet_by_name_mut("Sheet1")
            .unwrap()
            .get_cell_mut((1 as u32, (2 * i) as u32 + 2))
            .set_value(raws[i].clone().as_str());
    }

    let _ = umya_spreadsheet::writer::xlsx::write(&book, path);
}

pub fn out_one_row(raws1: Vec<String>, raws2: Vec<String>, book: &mut Spreadsheet) {
    let uuid = Uuid::new_v4().to_string();
    let path = format!("{}.xlsx", uuid);
    let path = std::path::Path::new(&path);

    for i in 0..raws1.len() {
        book.get_sheet_by_name_mut("Sheet1")
            .unwrap()
            .get_cell_mut((1 as u32, (i + 1) as u32))
            .set_value(raws1[i].clone().as_str());
    }

    for i in 0..raws2.len() {
        book.get_sheet_by_name_mut("Sheet1")
            .unwrap()
            .get_cell_mut((2 as u32, (i + 1) as u32))
            .set_value(raws2[i].clone().as_str());
    }

    let _ = umya_spreadsheet::writer::xlsx::write(&book, path);
}

pub fn read_row_from_file(path: &Path, column_idx: u32) -> Vec<String> {
    let book = umya_spreadsheet::reader::xlsx::read(path).unwrap();
    let sheet = book.get_sheet_by_name("Sheet1").unwrap();
    let raw_cells = sheet.get_collection_by_column(&column_idx);
    let res = raw_cells.into_iter().map(|raw| raw.get_value());
    println!("{:?}", res);
    vec!["".to_string()]
    // res
}

pub fn read_line_from_file(path: &Path) -> Vec<String> {
    let mut file = File::open(path).unwrap();
    let mut content = String::new();
    file.read_to_string(&mut content).unwrap();
    let strs = content
        .split("\n")
        .map(|str| str.trim().replace(" ", "").to_string())
        .collect::<Vec<String>>();
    return strs;
}
