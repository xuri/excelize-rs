// Copyright 2021 - 2024 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

use quick_xml::de::from_str;

use crate::{xml_rels, ExcelizeError, Spreadsheet};

pub trait Rels {
    fn rels_reader(&mut self, path: &str) -> Result<(), ExcelizeError>
    where
        Self: std::marker::Sized;
}

impl Rels for Spreadsheet {
    fn rels_reader(&mut self, path: &str) -> Result<(), ExcelizeError> {
        let rels: xml_rels::XMLRelationships;
        match self.rels.get_key_value(path) {
            Some(_) => Ok(()),
            None => {
                if let Some(buf) = self.file.get_key_value(path) {
                    let s = match std::str::from_utf8(&buf.1) {
                        Ok(v) => v,
                        Err(e) => return Err(ExcelizeError::CommonError(e.to_string())),
                    };
                    match from_str(s) {
                        Ok(o) => rels = o,
                        Err(e) => return Err(ExcelizeError::CommonError(e.to_string())),
                    }

                    self.rels.insert(String::from(path), rels);
                    return Ok(());
                }
                Err(ExcelizeError::CommonError(String::from("rels is none")))
            }
        }
    }
}
