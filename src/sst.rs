// Copyright 2021 The excelize-rs Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

use quick_xml::de::from_str;

use crate::{xml_sst, ExcelizeError, Spreadsheet};
pub trait SST {
    fn get_sst(&mut self) -> Result<(), ExcelizeError>
    where
        Self: std::marker::Sized;
}

impl SST for Spreadsheet {
    fn get_sst(&mut self) -> Result<(), ExcelizeError> {
        let sst: xml_sst::CTSST;
        match self.sst {
            Some(_) => Ok(()),
            None => {
                if let Some(buf) = self.file.get_key_value("xl/sharedStrings.xml") {
                    let s = match std::str::from_utf8(&buf.1) {
                        Ok(v) => v,
                        Err(e) => {
                            return Err(ExcelizeError::CommonError(e.to_string()));
                        }
                    };
                    match from_str(s) {
                        Ok(o) => sst = o,
                        Err(e) => {
                            return Err(ExcelizeError::CommonError(e.to_string()));
                        }
                    }
                    self.sst.replace(sst);
                    return Ok(());
                }
                Err(ExcelizeError::CommonError(String::from("sst is none")))
            }
        }
    }
}
