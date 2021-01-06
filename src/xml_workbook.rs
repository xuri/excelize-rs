// Copyright 2021 The excelize-rs Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

use serde::Deserialize;
use serde::Serialize;

/// XMLWorkbook contains elements and attributes that encompass the data
/// content of the workbook. The workbook's child elements each have their own
/// subclause references.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "workbook")]
pub struct XMLWorkbook {
    #[serde(rename = "fileVersion")]
    pub file_version: XMLFileVersion,
    #[serde(rename = "sheets")]
    pub sheets: XMLSheets,
}

/// XMLFileVersion directly maps the fileVersion element. This element defines
/// properties that track which version of the application accessed the data and
/// source code contained in the file.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "fileVersion")]
pub struct XMLFileVersion {
    #[serde(rename = "appName")]
    pub app_name: String,
    #[serde(rename = "lastEdited")]
    pub last_edited: i32,
    #[serde(rename = "lowestEdited")]
    pub lowest_edited: i32,
    #[serde(rename = "rupBuild")]
    pub rup_build: i32,
}

/// XMLSheets directly maps the sheets element from the namespace
/// http://schemas.openxmlformats.org/spreadsheetml/2006/main.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct XMLSheets {
    #[serde(rename = "sheet")]
    pub sheet: Vec<XMLSheet>,
}

/// XMLSheets defines a sheet in this workbook. Sheet data is stored in a
/// separate part.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct XMLSheet {
    #[serde(rename = "name")]
    pub name: String,
    #[serde(rename = "sheetId")]
    pub sheet_id: i32,
    #[serde(rename = "r:id")]
    pub id: String,
}
