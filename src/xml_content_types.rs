// Copyright 2021 - 2024 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

use serde::Deserialize;
use serde::Serialize;

/// XMLTypes directly maps the types element of content types for relationship
/// parts, it takes a Multipurpose Internet Mail Extension (MIME) media type as a
/// value.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "Types")]
pub struct XMLTypes {
    #[serde(rename = "$value")]
    fields: Vec<XMLOverride>,
}

/// CTMaps directly maps the elements in the namespace
/// http://schemas.openxmlformats.org/package/2006/content-types
#[derive(Debug, Deserialize, Serialize, PartialEq)]
pub enum XMLOverride {
    Override {
        #[serde(rename = "PartName")]
        part_name: String,
        #[serde(rename = "ContentType")]
        content_type: String,
    },
    Default {
        #[serde(rename = "Extension")]
        extension: String,
        #[serde(rename = "ContentType")]
        content_type: String,
    },
}
