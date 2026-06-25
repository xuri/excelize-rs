use crate::{xml_styles, ExcelizeError, Spreadsheet};
use quick_xml::de::from_str;

pub trait Styles {
    fn get_styles(&mut self) -> Result<(), ExcelizeError>;
}

impl Styles for Spreadsheet {
    fn get_styles(&mut self) -> Result<(), ExcelizeError> {
        if let Some(buf) = self.file.get_key_value("xl/styles.xml") {
            let s = std::str::from_utf8(&buf.1)
                .map_err(|e| ExcelizeError::CommonError(e.to_string()))?;
            let styles: xml_styles::CTStylesheet =
                from_str(s).map_err(|e| ExcelizeError::CommonError(e.to_string()))?;
            self.styles.replace(styles);
        }
        Ok(())
    }
}