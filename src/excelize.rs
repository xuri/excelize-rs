// Copyright 2021 The excelize-rs Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

use crate::{
    xml_content_types, xml_rels, xml_sst, xml_workbook, xml_worksheet, ExcelizeError, Rels,
    Workbook, Worksheet, SST,
};
use std::{collections::HashMap, fs, io};
use zip::ZipArchive;
#[derive(Debug)]
pub struct Spreadsheet {
    pub file: HashMap<String, Vec<u8>>,
    pub content_type: Option<xml_content_types::XMLTypes>,
    pub workbook: Option<xml_workbook::XMLWorkbook>,
    pub worksheets: HashMap<String, xml_worksheet::XMLWorksheet>,
    pub sst: Option<xml_sst::CTSST>,
    pub rels: HashMap<String, xml_rels::XMLRelationships>,
}

impl Spreadsheet {
    pub fn open_file(path: String) -> Result<Self, ExcelizeError> {
        let mut spreadsheet = Spreadsheet {
            file: HashMap::new(),
            content_type: None,
            workbook: None,
            worksheets: HashMap::new(),
            sst: None,
            rels: HashMap::new(),
        };
        match fs::File::open(&path) {
            Ok(file) => {
                let mut archive: zip::ZipArchive<std::fs::File>;
                match ZipArchive::new(file) {
                    Ok(content) => archive = content,
                    Err(_e) => {
                        return Err(ExcelizeError::CommonError(String::from("read zip error")))
                    }
                }
                for i in 0..archive.len() {
                    let mut file;
                    match archive.by_index(i) {
                        Ok(zip) => {
                            file = zip;
                        }
                        Err(e) => return Err(ExcelizeError::CommonError(e.to_string())),
                    }
                    let outpath = match file.enclosed_name() {
                        Some(path) => path.to_owned(),
                        None => continue,
                    };
                    if !(&*file.name()).ends_with('/') {
                        let mut data = vec![];
                        match io::copy(&mut file, &mut data) {
                            Ok(_) => (),
                            Err(e) => return Err(ExcelizeError::CommonError(e.to_string())),
                        }
                        match outpath.to_str() {
                            Some(xml_path) => {
                                spreadsheet.file.insert(String::from(xml_path), data);
                            }
                            None => {
                                return Err(ExcelizeError::CommonError(String::from(
                                    "extract zip failed",
                                )))
                            }
                        }
                    }
                }
                match spreadsheet.get_content_type() {
                    Ok(_) => {}
                    Err(e) => return Err(e),
                }
                match spreadsheet.get_workbook() {
                    Ok(_) => {}
                    Err(e) => return Err(e),
                }
                match spreadsheet.rels_reader(&"xl/_rels/workbook.xml.rels") {
                    Ok(_) => {}
                    Err(e) => return Err(e),
                }
                match spreadsheet.worksheet_reader() {
                    Ok(_) => {}
                    Err(e) => return Err(e),
                }
                if spreadsheet.get_sst().is_ok() {}
                Ok(spreadsheet)
            }
            Err(e) => Err(ExcelizeError::CommonError(e.to_string())),
        }
    }
}
// cargo test -- --color always --nocapture
#[cfg(test)]
mod tests {
    use super::*;
    use crate::Cell;
    #[test]
    fn test_open_file() {
        let path = String::from("src/test/Book1.xlsx");
        let wb = Spreadsheet::open_file(path);
        match wb {
            Ok(ws) => match ws.get_cell_value("Sheet1", 2, 1) {
                Ok(c) => {
                    let cell = String::from(c);
                    println!("the value of cell A1 is: {}", cell)
                }
                Err(e) => {
                    println!("{:?}", e);
                    panic!(e);
                }
            },
            Err(e) => {
                print!("{:?}", e);
                panic!(e);
            }
        }
    }
}
