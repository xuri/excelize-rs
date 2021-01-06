// Copyright 2021 The excelize-rs Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

use serde::Deserialize;
use serde::Serialize;

/// XMLWorksheet directly maps the worksheet element in the namespace
/// http://schemas.openxmlformats.org/spreadsheetml/2006/main.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "worksheet")]
pub struct XMLWorksheet {
    #[serde(rename = "sheetPr")]
    pub sheet_pr: Option<CTSheetPr>,
    #[serde(rename = "dimension")]
    pub dimension: Option<CTSheetDimension>,
    #[serde(rename = "sheetViews")]
    pub sheet_views: Option<CTSheetViews>,
    #[serde(rename = "sheetFormatPr")]
    pub sheet_format_pr: Option<CTSheetFormatPr>,
    #[serde(rename = "cols")]
    pub cols: Option<CTCols>,
    #[serde(rename = "sheetData")]
    pub sheet_data: CTSheetData,
    #[serde(rename = "sheetProtection")]
    pub sheet_protection: Option<CTSheetProtection>,
    #[serde(rename = "autoFilter")]
    pub auto_filter: Option<CTAutoFilter>,
    #[serde(rename = "sortState")]
    pub sort_state: Option<CTSortState>,
}

/// CTCols defines column width and column formatting for one or more columns
/// of the worksheet.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTCols {
    #[serde(rename = "col")]
    pub col: Vec<CTCol>,
}

/// CTCol directly maps the col (Column Width & Formatting). Defines column
/// width and column formatting for one or more columns of the worksheet.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTCol {
    #[serde(rename = "min")]
    pub min: u32,
    #[serde(rename = "max")]
    pub max: u32,
    #[serde(rename = "width")]
    pub width: Option<f64>,
    #[serde(rename = "style")]
    pub style: Option<u32>,
    #[serde(rename = "hidden")]
    pub hidden: Option<bool>,
    #[serde(rename = "bestFit")]
    pub best_fit: Option<bool>,
    #[serde(rename = "customWidth")]
    pub custom_width: Option<bool>,
    #[serde(rename = "phonetic")]
    pub phonetic: Option<bool>,
    #[serde(rename = "outlineLevel")]
    pub outline_level: Option<u8>,
    #[serde(rename = "collapsed")]
    pub collapsed: Option<bool>,
}

/// CTSheetDimension directly maps the dimension element in the namespace
/// http://schemas.openxmlformats.org/spreadsheetml/2006/main - This element
/// specifies the used range of the worksheet. It specifies the row and column
/// bounds of used cells in the worksheet. This is optional and is not
/// required. Used cells include cells with formulas, text content, and cell
/// formatting. When an entire column is formatted, only the first cell in that
/// column is considered used.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "dimension")]
pub struct CTSheetDimension {
    #[serde(rename = "ref")]
    pub ref_attr: String,
}

/// CTSheetViews represents worksheet views collection.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTSheetViews {
    #[serde(rename = "sheetView")]
    pub sheet_view: Vec<CTSheetView>,
    #[serde(rename = "extLst")]
    pub ext_lst: Option<String>,
}

/// CTSheetView represents a single sheet view definition. When more than one
/// sheet view is defined in the file, it means that when opening the workbook,
/// each sheet view corresponds to a separate window within the spreadsheet
/// application, where each window is showing the particular sheet containing
/// the same workbookViewId value, the last sheetView definition is loaded, and
/// the others are discarded. When multiple windows are viewing the same sheet,
/// multiple sheetView elements (with corresponding workbookView entries) are
/// saved.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTSheetView {
    #[serde(rename = "tabSelected")]
    pub tab_selected: Option<bool>,
    #[serde(rename = "workbookViewId")]
    pub workbook_view_id: i32,
}

/// CTSheetFormatPr directly maps the sheetFormatPr element in the namespace
/// http://schemas.openxmlformats.org/spreadsheetml/2006/main. This element
/// specifies the sheet formatting properties.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "sheetFormatPr")]
pub struct CTSheetFormatPr {
    #[serde(rename = "baseColWidth")]
    pub base_col_width: Option<u32>,
    #[serde(rename = "defaultColWidth")]
    pub default_col_width: Option<f64>,
    #[serde(rename = "defaultRowHeight")]
    pub default_row_height: f64,
    #[serde(rename = "customHeight")]
    pub custom_height: Option<bool>,
    #[serde(rename = "zeroHeight")]
    pub zero_height: Option<bool>,
    #[serde(rename = "thickTop")]
    pub thick_top: Option<bool>,
    #[serde(rename = "thickBottom")]
    pub thick_bottom: Option<bool>,
    #[serde(rename = "outlineLevelRow")]
    pub outline_level_row: Option<u8>,
    #[serde(rename = "outlineLevelCol")]
    pub outline_level_col: Option<u8>,
}

/// CTSheetData collection represents the cell table itself. This collection
/// expresses information about each cell, grouped together by rows in the
/// worksheet.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "sheetData")]
pub struct CTSheetData {
    #[serde(rename = "row")]
    pub row: Option<Vec<CTRow>>,
}

/// CTRow directly maps the row element. The element expresses information
/// about an entire row of a worksheet, and contains all cell definitions for a
/// particular row in the worksheet.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "row")]
pub struct CTRow {
    pub r: Option<u32>,
    #[serde(rename = "spans")]
    pub spans: Option<String>,
    #[serde(rename = "s")]
    pub s: Option<u32>,
    #[serde(rename = "customFormat")]
    pub custom_format: Option<bool>,
    #[serde(rename = "ht")]
    pub ht: Option<f64>,
    #[serde(rename = "hidden")]
    pub hidden: Option<bool>,
    #[serde(rename = "customHeight")]
    pub custom_height: Option<bool>,
    #[serde(rename = "outlineLevel")]
    pub outline_level: Option<u8>,
    #[serde(rename = "collapsed")]
    pub collapsed: Option<bool>,
    #[serde(rename = "thickTop")]
    pub thick_top: Option<bool>,
    #[serde(rename = "thickBot")]
    pub thick_bot: Option<bool>,
    #[serde(rename = "ph")]
    pub ph: Option<bool>,
    #[serde(rename = "c")]
    pub c: Vec<CTCell>,
    #[serde(rename = "extLst")]
    pub ext_lst: Option<String>,
}

/// CTCell collection represents a cell in the worksheet. Information about the
/// cell's location (reference), value, data type, formatting, and formula is
/// expressed here.
///
/// This simple type is restricted to the values listed in the following table:
///
///      Enumeration Value         | Description
///     ---------------------------+---------------------------------
///      b (Boolean)               | Cell containing a boolean.
///      d (Date)                  | Cell contains a date in the ISO 8601 format.
///      e (Error)                 | Cell containing an error.
///      inlineStr (Inline String) | Cell containing an (inline) rich string, i.e., one not in the shared string table. If this cell type is used, then the cell value is in the is element rather than the v element in the cell (c element).
///      n (Number)                | Cell containing a number.
///      s (Shared String)         | Cell containing a shared string.
///      str (String)              | Cell containing a formula string.
///
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "c")]
pub struct CTCell {
    #[serde(rename = "r")]
    pub r: String,
    #[serde(rename = "t")]
    pub t: Option<String>,
    #[serde(rename = "v")]
    pub v: Option<String>,
}

/// CTSheetPr directly maps the sheetPr element in the namespace
/// http://schemas.openxmlformats.org/spreadsheetml/2006/main - Sheet-level
/// properties.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTSheetPr {
    #[serde(rename = "syncHorizontal")]
    pub sync_horizontal: Option<bool>,
    #[serde(rename = "syncVertical")]
    pub sync_vertical: Option<bool>,
    #[serde(rename = "syncRef")]
    pub sync_ref: Option<String>,
    #[serde(rename = "transitionEvaluation")]
    pub transition_evaluation: Option<bool>,
    #[serde(rename = "transitionEntry")]
    pub transition_entry: Option<bool>,
    #[serde(rename = "published")]
    pub published: Option<bool>,
    #[serde(rename = "codeName")]
    pub code_name: Option<String>,
    #[serde(rename = "filterMode")]
    pub filter_mode: Option<bool>,
    #[serde(rename = "enableFormatConditionsCalculation")]
    pub enable_format_conditions_calculation: Option<bool>,
    #[serde(rename = "tabColor")]
    pub tab_color: Vec<CTColor>,
    #[serde(rename = "outlinePr")]
    pub outline_pr: Vec<CTOutlinePr>,
    #[serde(rename = "pageSetUpPr")]
    pub page_set_up_pr: Vec<CTPageSetUpPr>,
}

/// CTColor is a common mapping used for both the fgColor and bgColor elements.
/// Foreground color of the cell fill pattern. Cell fill patterns operate with
/// two colors: a background color and a foreground color. These combine together
/// to make a patterned cell fill. Background color of the cell fill pattern.
/// Cell fill patterns operate with two colors: a background color and a
/// foreground color. These combine together to make a patterned cell fill.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTColor {
    #[serde(rename = "auto")]
    pub auto: Option<bool>,
    #[serde(rename = "indexed")]
    pub indexed: Option<u32>,
    #[serde(rename = "rgb")]
    pub rgb: Option<String>,
    #[serde(rename = "theme")]
    pub theme: Option<u32>,
    #[serde(rename = "tint")]
    pub tint: Option<f64>,
}

/// CTOutlinePr maps to the outlinePr element. SummaryBelow allows you to
/// adjust the direction of grouper controls.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTOutlinePr {
    #[serde(rename = "applyStyles")]
    pub apply_styles: Option<bool>,
    #[serde(rename = "summaryBelow")]
    pub summary_below: Option<bool>,
    #[serde(rename = "summaryRight")]
    pub summary_right: Option<bool>,
    #[serde(rename = "showOutlineSymbols")]
    pub show_outline_symbols: Option<bool>,
}

/// CTPageSetUpPr expresses page setup properties of the worksheet.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTPageSetUpPr {
    #[serde(rename = "autoPageBreaks")]
    pub auto_page_breaks: Option<bool>,
    #[serde(rename = "fitToPage")]
    pub fit_to_page: Option<bool>,
}

/// CTSheetProtection collection expresses the sheet protection options to
/// enforce when the sheet is protected.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTSheetProtection {
    #[serde(rename = "algorithmName")]
    pub algorithm_name: Option<String>,
    #[serde(rename = "hashValue")]
    pub hash_value: Option<String>,
    #[serde(rename = "saltValue")]
    pub salt_value: Option<String>,
    #[serde(rename = "spinCount")]
    pub spin_count: Option<u32>,
    #[serde(rename = "sheet")]
    pub sheet: Option<bool>,
    #[serde(rename = "objects")]
    pub objects: Option<bool>,
    #[serde(rename = "scenarios")]
    pub scenarios: Option<bool>,
    #[serde(rename = "formatCells")]
    pub format_cells: Option<bool>,
    #[serde(rename = "formatColumns")]
    pub format_columns: Option<bool>,
    #[serde(rename = "formatRows")]
    pub format_rows: Option<bool>,
    #[serde(rename = "insertColumns")]
    pub insert_columns: Option<bool>,
    #[serde(rename = "insertRows")]
    pub insert_rows: Option<bool>,
    #[serde(rename = "insertHyperlinks")]
    pub insert_hyperlinks: Option<bool>,
    #[serde(rename = "deleteColumns")]
    pub delete_columns: Option<bool>,
    #[serde(rename = "deleteRows")]
    pub delete_rows: Option<bool>,
    #[serde(rename = "selectLockedCells")]
    pub select_locked_cells: Option<bool>,
    #[serde(rename = "sort")]
    pub sort: Option<bool>,
    #[serde(rename = "autoFilter")]
    pub auto_filter: Option<bool>,
    #[serde(rename = "pivotTables")]
    pub pivot_tables: Option<bool>,
    #[serde(rename = "selectUnlockedCells")]
    pub select_unlocked_cells: Option<bool>,
}

/// CTAutoFilter temporarily hides rows based on a filter criteria, which is
/// applied column by column to a table of data in the worksheet. This collection
/// expresses AutoFilter settings.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTAutoFilter {
    #[serde(rename = "ref")]
    pub ref_attr: Option<String>,
    #[serde(rename = "filterColumn")]
    pub filter_column: Vec<CTFilterColumn>,
    #[serde(rename = "sortState")]
    pub sort_state: Option<CTSortState>,
}

/// CTFilterColumn directly maps the filterColumn element. The filterColumn
/// collection identifies a particular column in the AutoFilter range and
/// specifies filter information that has been applied to this column. If a
/// column in the AutoFilter range has no criteria specified, then there is no
/// corresponding filterColumn collection expressed for that column.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTFilterColumn {
    #[serde(rename = "colId")]
    pub col_id: u32,
    #[serde(rename = "hiddenButton")]
    pub hidden_button: Option<bool>,
    #[serde(rename = "showButton")]
    pub show_button: Option<bool>,
    #[serde(rename = "filters")]
    pub filters: Option<CTFilters>,
    #[serde(rename = "top10")]
    pub top10: Option<CTTop10>,
    #[serde(rename = "customFilters")]
    pub custom_filters: Option<CTCustomFilters>,
    #[serde(rename = "dynamicFilter")]
    pub dynamic_filter: Option<CTDynamicFilter>,
    #[serde(rename = "colorFilter")]
    pub color_filter: Option<CTColorFilter>,
    #[serde(rename = "iconFilter")]
    pub icon_filter: Option<CTIconFilter>,
    #[serde(rename = "extLst")]
    pub ext_lst: Option<String>,
}
/// CTColorFilter directly maps the colorFilter element. This element specifies
/// the color to filter by and whether to use the cell's fill or font color in
/// the filter criteria. If the cell's font or fill color does not match the
/// color specified in the criteria, the rows corresponding to those cells are
/// hidden from view.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTColorFilter {
    #[serde(rename = "dxfId")]
    pub dxf_id: Option<u32>,
    #[serde(rename = "cellColor")]
    pub cell_color: Option<bool>,
}

/// CTIconFilter directly maps the iconFilter element. This element specifies
/// the icon set and particular icon within that set to filter by. For any cells
/// whose icon does not match the specified criteria, the corresponding rows
/// shall be hidden from view when the filter is applied.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTIconFilter {
    #[serde(rename = "iconSet")]
    pub icon_set: String,
    #[serde(rename = "iconId")]
    pub icon_id: Option<u32>,
}

/// CTDynamicFilter directly maps the dynamicFilter element. This collection
/// specifies dynamic filter criteria. These criteria are considered dynamic
/// because they can change, either with the data itself (e.g., "above average")
/// or with the current system date (e.g., show values for "today"). For any
/// cells whose values do not meet the specified criteria, the corresponding rows
/// shall be hidden from view when the filter is applied.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTDynamicFilter {
    #[serde(rename = "type")]
    pub type_attr: String,
    #[serde(rename = "val")]
    pub val: Option<f64>,
    #[serde(rename = "valIso")]
    pub val_iso: Option<u8>,
    #[serde(rename = "maxValIso")]
    pub max_val_iso: Option<u8>,
}

/// CTCustomFilters directly maps the customFilters element. When there is more
/// than one custom filter criteria to apply (an 'and' or 'or' joining two
/// criteria), then this element groups the customFilter elements together.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTCustomFilters {
    #[serde(rename = "and")]
    pub and: Option<bool>,
    #[serde(rename = "customFilter")]
    pub custom_filter: Vec<CTCustomFilter>,
}

/// CTCustomFilter directly maps the customFilter element. A custom AutoFilter
/// specifies an operator and a value. There can be at most two customFilters
/// specified, and in that case the parent element specifies whether the two
/// conditions are joined by 'and' or 'or'. For any cells whose values do not
/// meet the specified criteria, the corresponding rows shall be hidden from view
/// when the filter is applied.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTCustomFilter {
    #[serde(rename = "operator")]
    pub operator: Option<String>,
    #[serde(rename = "val")]
    pub val: Option<String>,
}

/// CTTop10 directly maps the top10 element. This element specifies the top N
/// (percent or number of items) to filter by.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTTop10 {
    #[serde(rename = "top")]
    pub top: Option<bool>,
    #[serde(rename = "percent")]
    pub percent: Option<bool>,
    #[serde(rename = "val")]
    pub val: f64,
    #[serde(rename = "filterVal")]
    pub filter_val: Option<f64>,
}

/// CTFilters directly maps the filters (Filter Criteria) element. When
/// multiple values are chosen to filter by, or when a group of date values are
/// chosen to filter by, this element groups those criteria together.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTFilters {
    #[serde(rename = "blank")]
    pub blank: Option<bool>,
    #[serde(rename = "calendarType")]
    pub calendar_type: Option<String>,
    #[serde(rename = "filter")]
    pub filter: Vec<CTFilter>,
    #[serde(rename = "dateGroupItem")]
    pub date_group_item: Vec<CTDateGroupItem>,
}

/// CTFilter directly maps the filter element. This element expresses a filter
/// criteria value.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTFilter {
    #[serde(rename = "val")]
    pub val: Option<String>,
}

/// CTDateGroupItem directly maps the dateGroupItem element. This collection is
/// used to express a group of dates or times which are used in an AutoFilter
/// criteria. [Note: See parent element for an example. end note] Values are
/// always written in the calendar type of the first date encountered in the
/// filter range, so that all subsequent dates, even when formatted or
/// represented by other calendar types, can be correctly compared for the
/// purposes of filtering.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTDateGroupItem {
    #[serde(rename = "year")]
    pub year: u16,
    #[serde(rename = "month")]
    pub month: Option<u16>,
    #[serde(rename = "day")]
    pub day: Option<u16>,
    #[serde(rename = "hour")]
    pub hour: Option<u16>,
    #[serde(rename = "minute")]
    pub minute: Option<u16>,
    #[serde(rename = "second")]
    pub second: Option<u16>,
    #[serde(rename = "dateTimeGrouping")]
    pub date_time_grouping: String,
}

/// CTSortState directly maps the sortState element. This collection
/// preserves the AutoFilter sort state.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTSortState {
    #[serde(rename = "columnSort")]
    pub column_sort: Option<bool>,
    #[serde(rename = "caseSensitive")]
    pub case_sensitive: Option<bool>,
    #[serde(rename = "sortMethod")]
    pub sort_method: Option<String>,
    #[serde(rename = "ref")]
    pub ref_attr: String,
    #[serde(rename = "sortCondition")]
    pub sort_condition: Vec<CTSortCondition>,
    #[serde(rename = "extLst")]
    pub ext_lst: Option<String>,
}

#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub struct CTSortCondition {
    #[serde(rename = "descending")]
    pub descending: Option<bool>,
    #[serde(rename = "sortBy")]
    pub sort_by: Option<String>,
    #[serde(rename = "ref")]
    pub ref_attr: String,
    #[serde(rename = "customList")]
    pub custom_list: Option<String>,
    #[serde(rename = "dxfId")]
    pub dxf_id: Option<u32>,
    #[serde(rename = "iconSet")]
    pub icon_set: Option<String>,
    #[serde(rename = "iconId")]
    pub icon_id: Option<u32>,
}
