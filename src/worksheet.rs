// Copyright 2021 - 2024 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

use crate::{xml_worksheet, ExcelizeError, Spreadsheet};
use std::str;
extern crate quick_xml;
extern crate serde;

use quick_xml::de::from_str;
pub trait Worksheet {
    fn worksheet_reader(&mut self) -> Result<(), ExcelizeError>
    where
        Self: std::marker::Sized;
    fn get_target_by_rid(&self, rid: String) -> Result<String, ExcelizeError>;
}

impl Worksheet for Spreadsheet {
    fn worksheet_reader(&mut self) -> Result<(), ExcelizeError> {
        match self.workbook {
            Some(ref wb) => {
                for sheet in &wb.sheets.sheet {
                    let ws: xml_worksheet::XMLWorksheet;
                    match self.get_target_by_rid(String::from(&sheet.id)) {
                        Ok(ref target) => match self.worksheets.get_key_value(&sheet.name) {
                            Some(_) => {
                                return Ok(());
                            }
                            None => {
                                let rels_xml = format!("xl/{}", &target.to_string());
                                match self.file.get_key_value(&rels_xml.to_string()) {
                                    Some(buf) => {
                                        let s = match str::from_utf8(&buf.1) {
                                            Ok(v) => v,
                                            Err(e) => {
                                                return Err(ExcelizeError::CommonError(
                                                    e.to_string(),
                                                ))
                                            }
                                        };
                                        match from_str(s) {
                                            Ok(o) => ws = o,
                                            Err(e) => {
                                                return Err(ExcelizeError::CommonError(
                                                    e.to_string(),
                                                ))
                                            }
                                        }

                                        self.worksheets.insert(sheet.name.to_string(), ws);
                                        continue;
                                    }
                                    None => {
                                        return Err(ExcelizeError::CommonError(String::from(
                                            "sheet is none",
                                        )));
                                    }
                                }
                            }
                        },
                        Err(e) => return Err(e),
                    }
                }
            }
            None => {
                return Err(ExcelizeError::CommonError(String::from("sheet is none")));
            }
        }
        Ok(())
    }

    fn get_target_by_rid(&self, rid: String) -> Result<String, ExcelizeError> {
        match self.rels.get_key_value("xl/_rels/workbook.xml.rels") {
            Some(rels) => {
                let mut target = String::from("");
                for rel in &rels.1.relationship {
                    if rel.id == rid {
                        target = String::from(&rel.target);
                        break;
                    }
                }
                Ok(target)
            }
            None => Err(ExcelizeError::CommonError(String::from("target is none"))),
        }
    }
}
