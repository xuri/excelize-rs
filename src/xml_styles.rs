// Copyright 2021 - 2024 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

use serde::Deserialize;
use serde::Serialize;

use crate::{CTColor, STXBool, STXInt, STXstring};

#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "styleSheet")]
pub struct CTStylesheet {
    #[serde(rename = "numFmts")]
    pub num_fmts: Option<CTNumFmts>,
    #[serde(rename = "fonts")]
    pub fonts: CTFonts,
    #[serde(rename = "fills")]
    pub fills: CTFills,
    #[serde(rename = "borders")]
    pub borders: CTBorders,
    #[serde(rename = "cellStyleXfs")]
    pub cell_style_xfs: Option<CTCellStyleXfs>,
    #[serde(rename = "cellXfs")]
    pub cell_xfs: CTCellXfs,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTNumFmts {
    #[serde(rename = "numFmt")]
    pub num_fmt: Vec<CTNumFmt>,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTNumFmt {
    #[serde(rename = "numFmtId")]
    pub num_fmt_id: u32,
    #[serde(rename = "formatCode")]
    pub format_code: String,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTFonts {
    #[serde(rename = "font")]
    pub font: Vec<CTFont>,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTFont {
    #[serde(rename = "name")]
    pub name: Option<STXstring>,
    #[serde(rename = "sz")]
    pub sz: Option<STXInt>,
    #[serde(rename = "color")]
    pub color: Option<CTColor>,
    #[serde(rename = "b")]
    pub b: Option<STXBool>,
    #[serde(rename = "i")]
    pub i: Option<STXBool>,
    #[serde(rename = "u")]
    pub u: Option<STXBool>,
    #[serde(rename = "strike")]
    pub strike: Option<STXBool>,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTFills {
    #[serde(rename = "fill")]
    pub fill: Vec<CTFill>,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTFill {
    #[serde(rename = "patternFill")]
    pub pattern_fill: Option<CTPatternFill>,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTPatternFill {
    #[serde(rename = "patternType")]
    pub pattern_type: Option<String>,
    #[serde(rename = "fgColor")]
    pub fg_color: Option<CTColor>,
    #[serde(rename = "bgColor")]
    pub bg_color: Option<CTColor>,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTBorders {
    #[serde(rename = "border")]
    pub border: Vec<CTBorder>,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTBorder {
    #[serde(rename = "left")]
    pub left: Option<CTBorderPr>,
    #[serde(rename = "right")]
    pub right: Option<CTBorderPr>,
    #[serde(rename = "top")]
    pub top: Option<CTBorderPr>,
    #[serde(rename = "bottom")]
    pub bottom: Option<CTBorderPr>,
    #[serde(rename = "diagonal")]
    pub diagonal: Option<CTBorderPr>,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTBorderPr {
    #[serde(rename = "style")]
    pub style: Option<String>,
    #[serde(rename = "color")]
    pub color: Option<CTColor>,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTCellStyleXfs {
    #[serde(rename = "xf")]
    pub xf: Vec<CTXf>,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTCellXfs {
    #[serde(rename = "xf")]
    pub xf: Vec<CTXf>,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTXf {
    #[serde(rename = "numFmtId")]
    pub num_fmt_id: Option<u32>,
    #[serde(rename = "fontId")]
    pub font_id: Option<u32>,
    #[serde(rename = "fillId")]
    pub fill_id: Option<u32>,
    #[serde(rename = "borderId")]
    pub border_id: Option<u32>,
}