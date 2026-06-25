// Copyright 2021 - 2024 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

use crate::{column_number_to_name, CTCell, ExcelizeError, Spreadsheet};

pub trait Cell {
    /// GetCellValue provides a function to get formatted value from cell by given
    /// worksheet name and axis in spreadsheet.
    fn get_cell_value(&self, sheet: &str, row: u32, col: u32) -> Result<String, ExcelizeError>;
    fn get_value_from(&self, cell: &CTCell) -> String;
    /// GetCellStyle provides a function to get the resolved style (numFmt/font/fill/border
    /// ids) of a cell by given worksheet name and axis in spreadsheet.
    fn get_cell_style(&self, sheet: &str, row: u32, col: u32) -> Result<CellStyle, ExcelizeError>;
    fn find_cell(&self, sheet: &str, row: u32, col: u32) -> Result<Option<&CTCell>, ExcelizeError>;
}

#[derive(Debug, Default, PartialEq)]
pub struct CellStyle {
    pub num_fmt_id: Option<u32>,
    pub font_id: Option<u32>,
    pub fill_id: Option<u32>,
    pub border_id: Option<u32>,
}

impl Cell for Spreadsheet {
    fn find_cell(&self, sheet: &str, row: u32, col: u32) -> Result<Option<&CTCell>, ExcelizeError> {
        let column_title = column_number_to_name(col)?;
        let target_ref = format!("{}{}", column_title, row);

        match self.worksheets.get_key_value(sheet) {
            Some(ws) => match &ws.1.sheet_data.row {
                Some(xml_row) => {
                    for r in xml_row {
                        if r.r == Some(row) {
                            for c in &r.c {
                                if c.r == target_ref {
                                    return Ok(Some(c));
                                }
                            }
                        }
                    }
                    Ok(None)
                }
                None => Ok(None),
            },
            None => Err(ExcelizeError::CommonError(format!(
                "sheet {} is not exist",
                sheet
            ))),
        }
    }

    fn get_cell_value(&self, sheet: &str, row: u32, col: u32) -> Result<String, ExcelizeError> {
        match self.find_cell(sheet, row, col)? {
            Some(c) => Ok(self.get_value_from(c)),
            None => Ok(String::from("")),
        }
    }

    fn get_cell_style(&self, sheet: &str, row: u32, col: u32) -> Result<CellStyle, ExcelizeError> {
        let style_idx = match self.find_cell(sheet, row, col)? {
            Some(c) => c.s.unwrap_or(0),
            None => 0,
        };

        match &self.styles {
            Some(styles) => match styles.cell_xfs.xf.get(style_idx as usize) {
                Some(xf) => Ok(CellStyle {
                    num_fmt_id: xf.num_fmt_id,
                    font_id: xf.font_id,
                    fill_id: xf.fill_id,
                    border_id: xf.border_id,
                }),
                None => Ok(CellStyle::default()),
            },
            None => Err(ExcelizeError::CommonError(String::from("styles is none"))),
        }
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