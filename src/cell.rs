// Copyright 2021 - 2024 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

use crate::{column_number_to_name, CTCell, ExcelizeError, Spreadsheet};
pub trait Cell {
    /// GetCellValue provides a function to get formatted value from cell by given
    /// worksheet name and axis in spreadsheet.
    fn get_cell_value(&self, sheet: &str, row: u32, col: u32) -> Result<String, ExcelizeError>;
    fn get_value_from(&self, cell: &CTCell) -> String;
}

impl Cell for Spreadsheet {
    fn get_cell_value(&self, sheet: &str, row: u32, col: u32) -> Result<String, ExcelizeError> {
        let empty = String::from("");
        let column_title;
        match column_number_to_name(col) {
            Ok(c) => column_title = c,
            Err(e) => return Err(e),
        }
        let worksheet = self.worksheets.get_key_value(sheet);
        match worksheet {
            Some(ws) => match &ws.1.sheet_data.row {
                Some(xml_row) => {
                    for r in xml_row {
                        match r.r {
                            Some(rn) => {
                                if rn == row {
                                    for c in &r.c {
                                        if c.r == format!("{}{}", column_title, row.to_string()) {
                                            return Ok(self.get_value_from(&c));
                                        }
                                    }
                                }
                            }
                            None => return Ok(empty),
                        }
                    }
                }
                None => return Ok(empty),
            },
            None => {
                return Err(ExcelizeError::CommonError(format!(
                    "sheet {} is not exist",
                    sheet
                )))
            }
        }
        Ok(empty)
    }

    fn get_value_from(&self, cell: &CTCell) -> String {
        let empty = String::from("");
        match cell.t {
            Some(ref t) => match &t[..] {
                "s" => match self.sst {
                    None => empty,
                    Some(ref shared_string) => {
                        let i;
                        match &cell.v {
                            Some(ref v) => {
                                match v.to_string().parse::<usize>() {
                                    Ok(idx) => {
                                        i = idx;
                                    }
                                    Err(_) => return String::from(v),
                                }
                                let si = &shared_string.si[i];
                                match si.t {
                                    Some(ref t) => t[0].to_string(),
                                    None => match si.r {
                                        Some(ref relts) => {
                                            let mut v = String::from("");
                                            for relt in relts {
                                                // TODO: preserve xml:space
                                                v.push_str(&relt.t.to_string());
                                            }
                                            v
                                        }
                                        None => String::from(v),
                                    },
                                }
                            }
                            None => empty,
                        }
                    }
                },
                "str" => match &cell.v {
                    Some(ref v) => String::from(v),
                    None => empty,
                },
                _ => match &cell.v {
                    Some(ref v) => String::from(v),
                    None => empty,
                },
            },
            None => match &cell.v {
                Some(ref v) => String::from(v),
                None => empty,
            },
        }
    }
}
