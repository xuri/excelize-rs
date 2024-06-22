// Copyright 2021 - 2024 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

use crate::{xml_content_types, xml_workbook, ExcelizeError, Spreadsheet};
use std::str;
extern crate quick_xml;
extern crate serde;

use quick_xml::de::from_str;
pub trait Workbook {
    fn get_content_type(&mut self) -> Result<(), ExcelizeError>
    where
        Self: std::marker::Sized;
    fn get_workbook(&mut self) -> Result<(), ExcelizeError>
    where
        Self: std::marker::Sized;
}

impl Workbook for Spreadsheet {
    fn get_content_type(&mut self) -> Result<(), ExcelizeError> {
        let types: xml_content_types::XMLTypes;
        match self.content_type {
            Some(_) => Ok(()),
            None => {
                if let Some(buf) = self.file.get_key_value("[Content_Types].xml") {
                    let s = match str::from_utf8(&buf.1) {
                        Ok(v) => v,
                        Err(e) => return Err(ExcelizeError::CommonError(e.to_string())),
                    };
                    match from_str(s) {
                        Ok(o) => types = o,
                        Err(e) => return Err(ExcelizeError::CommonError(e.to_string())),
                    }
                    self.content_type.replace(types);
                    return Ok(());
                }
                Err(ExcelizeError::CommonError(String::from(
                    "content type is none",
                )))
            }
        }
    }

    fn get_workbook(&mut self) -> Result<(), ExcelizeError> {
        let workbook: xml_workbook::XMLWorkbook;
        match self.workbook {
            Some(_) => Ok(()),
            None => {
                if let Some(buf) = self.file.get_key_value("xl/workbook.xml") {
                    let s = match str::from_utf8(&buf.1) {
                        Ok(v) => v,
                        Err(e) => return Err(ExcelizeError::CommonError(e.to_string())),
                    };
                    match from_str(s) {
                        Ok(o) => workbook = o,
                        Err(e) => {
                            println!("{}", e);
                            return Err(ExcelizeError::CommonError(e.to_string()));
                        }
                    }
                    self.workbook.replace(workbook);
                    return Ok(());
                }
                Err(ExcelizeError::CommonError(String::from("workbook is none")))
            }
        }
    }
}
