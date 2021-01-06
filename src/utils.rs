// Copyright 2021 The excelize-rs Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

use crate::ExcelizeError;

static _MAX_FONT_FAMILY_LENGTH: u32 = 31;
static _MAX_FONT_SIZE: u32 = 409;
static _MAX_FILE_NAME_LENGTH: u32 = 207;
static _MAX_COLUMN_WIDTH: u32 = 255;
static _MAX_ROW_HEIGHT: u32 = 409;
static _TOTAL_ROWS: u32 = 1048576;
static TOTAL_COLUMNS: u32 = 16384;
static _TOTAL_SHEET_HYPERLINKS: u32 = 65529;
static _TOTAL_CELL_CHARS: u32 = 32767;

// column_number_to_name provides a function to convert the integer to Excel
// sheet column title.
pub fn column_number_to_name(col: u32) -> Result<String, ExcelizeError> {
    if col < 1 {
        let err = format!("incorrect column number {}", col);
        return Err(ExcelizeError::CommonError(err));
    }
    if col > TOTAL_COLUMNS {
        return Err(ExcelizeError::CommonError(String::from(
            "column number exceeds maximum limit",
        )));
    }
    let mut column = "".to_string();
    if col > 0 {
        let mut v = col;
        let a = b'A';
        while v > 0 {
            let curt_i = ((v - 1) % 26) as u8;
            let curt_c = (a + curt_i) as char;
            column.insert(0, curt_c);
            v = (v - 1) / 26;
        }
    }
    Ok(column)
}
