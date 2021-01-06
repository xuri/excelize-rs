// Copyright 2021 The excelize-rs Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

use serde::Deserialize;
use serde::Serialize;

/// Relationships describe references from parts to other internal resources in the package or to external resources.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "Relationships")]
pub struct XMLRelationships {
    #[serde(rename = "Relationship")]
    pub relationship: Vec<XMLRelationship>,
}

/// XMLRelationship contains relations which maps id and XML.
#[derive(Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename = "Relationship")]
pub struct XMLRelationship {
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(rename = "Type")]
    pub rel_type: String,
    #[serde(rename = "Target")]
    pub target: String,
}
