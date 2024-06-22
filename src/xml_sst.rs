// Copyright 2021 - 2024 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

use serde::Deserialize;
use serde::Serialize;

/// CTSST directly maps the sst element from the namespace
/// http://schemas.openxmlformats.org/spreadsheetml/2006/main. String values may
/// be stored directly inside spreadsheet cell elements; however, storing the
/// same value inside multiple cell elements can result in very large worksheet
/// Parts, possibly resulting in performance degradation. The Shared String Table
/// is an indexed list of string values, shared across the workbook, which allows
/// implementations to store values only once.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "sst")]
pub struct CTSST {
    #[serde(rename = "count")]
    pub count: Option<i32>,
    #[serde(rename = "uniqueCount")]
    pub unique_count: Option<i32>,
    #[serde(rename = "si")]
    pub si: Vec<CTRst>,
    #[serde(rename = "extLst")]
    pub ext_lst: Option<String>,
}

/// CTRst (String Item) is the representation of an individual string in the
/// Shared String table. If the string is just a simple string with formatting
/// applied at the cell level, then the String Item (si) should contain a
/// single text element used to express the string. However, if the string in
/// the cell is more complex - i.e., has formatting applied at the character
/// level - then the string item shall consist of multiple rich text runs which
/// collectively are used to express the string.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "si")]
pub struct CTRst {
    #[serde(rename = "t")]
    pub t: Option<Vec<String>>,
    #[serde(rename = "r")]
    pub r: Option<Vec<RElt>>,
    #[serde(rename = "rPh")]
    pub r_ph: Option<Vec<CTPhoneticRun>>,
    #[serde(rename = "phoneticPr")]
    pub phonetic_pr: Option<Vec<CTPhoneticPr>>,
}

/// CTPhoneticRun element represents a run of text which displays a phonetic
/// hint for this String Item (si). Phonetic hints are used to give information
/// about the pronunciation of an East Asian language. The hints are displayed
/// as text within the spreadsheet cells across the top portion of the cell.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "rPh")]
pub struct CTPhoneticRun {}

/// CTPhoneticPr (Phonetic Properties) represents a collection of phonetic
/// properties that affect the display of phonetic text for this String Item
/// (si). Phonetic text is used to give hints as to the pronunciation of an East
/// Asian language, and the hints are displayed as text within the spreadsheet
/// cells across the top portion of the cell. Since the phonetic hints are text,
/// every phonetic hint is expressed as a phonetic run (rPh), and these
/// properties specify how to display that phonetic run.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "phoneticPr")]
pub struct CTPhoneticPr {}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "r")]
pub struct RElt {
    #[serde(rename = "t")]
    pub t: String,
    #[serde(rename = "rPr")]
    pub r_pr: Option<Vec<RPrElt>>,
}

/// RPrElt (Run Properties) specifies a set of run properties which shall be
/// applied to the contents of the parent run after all style formatting has been
/// applied to the text. These properties are defined as direct formatting, since
/// they are directly applied to the run and supersede any formatting from
/// styles.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "rPr")]
pub struct RPrElt {
    #[serde(rename = "rFont")]
    pub r_font: Option<STXstring>,
    #[serde(rename = "charset")]
    pub charset: Option<STXInt>,
    #[serde(rename = "family")]
    pub family: Option<STXInt>,
    #[serde(rename = "b")]
    pub b: Option<STXBool>,
    #[serde(rename = "i")]
    pub i: Option<STXBool>,
    #[serde(rename = "strike")]
    pub strike: Option<STXBool>,
    #[serde(rename = "outline")]
    pub outline: Option<STXBool>,
    #[serde(rename = "shadow")]
    pub shadow: Option<STXBool>,
    #[serde(rename = "condense")]
    pub condense: Option<STXBool>,
    #[serde(rename = "extend")]
    pub extend: Option<STXBool>,
}

/// STXstring directly maps the val element with string data type as an
/// attribute
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct STXstring {
    #[serde(rename = "val")]
    pub val: String,
}

/// STXInt directly maps the val element with integer data type as an
/// attribute
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct STXInt {
    #[serde(rename = "val")]
    pub val: i32,
}

/// STXBool directly maps the val element with boolean data type as an
/// attribute
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct STXBool {
    #[serde(rename = "val")]
    pub val: Option<bool>,
}
