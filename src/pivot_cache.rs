// pivot_cache - A module for creating Excel PivotCache Definition and Records.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

mod tests;

use std::io::Cursor;

use crate::xmlwriter::{xml_declaration, xml_empty_tag, xml_end_tag, xml_start_tag};

// -----------------------------------------------------------------------
// Enums
// -----------------------------------------------------------------------

/// The source type for pivot cache data.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PivotCacheSourceType {
    /// Data comes from a worksheet range.
    #[default]
    Worksheet,

    /// Data comes from an external connection.
    External,

    /// Data comes from consolidation of multiple ranges.
    Consolidation,
}

impl PivotCacheSourceType {
    /// Convert to XML type attribute value.
    pub(crate) fn to_type_str(&self) -> &'static str {
        match self {
            PivotCacheSourceType::Worksheet => "worksheet",
            PivotCacheSourceType::External => "external",
            PivotCacheSourceType::Consolidation => "consolidation",
        }
    }
}

/// Represents a value type in the pivot cache.
#[derive(Debug, Clone, PartialEq)]
pub enum PivotCacheValue {
    /// String value.
    String(String),

    /// Numeric value.
    Number(f64),

    /// Boolean value.
    Boolean(bool),

    /// Date/time value (stored as string in ISO format).
    DateTime(String),

    /// Error value.
    Error(String),

    /// Missing/null value.
    Missing,
}

// -----------------------------------------------------------------------
// PivotCacheField
// -----------------------------------------------------------------------

/// Represents a field in the pivot cache.
///
/// Each cache field corresponds to a column in the source data and stores
/// unique values (shared items) for that field.
#[derive(Debug, Clone)]
pub struct PivotCacheField {
    /// Field name (column header).
    pub(crate) name: String,

    /// Number format ID.
    pub(crate) num_fmt_id: u32,

    /// Shared items (unique values).
    pub(crate) shared_items: Vec<PivotCacheValue>,

    /// Whether the field contains string values.
    pub(crate) contains_string: bool,

    /// Whether the field contains numeric values.
    pub(crate) contains_number: bool,

    /// Whether the field contains integer values.
    pub(crate) contains_integer: bool,

    /// Whether the field contains blank values.
    pub(crate) contains_blank: bool,

    /// Whether the field contains mixed types.
    pub(crate) contains_mixed_types: bool,

    /// Minimum numeric value.
    pub(crate) min_value: Option<f64>,

    /// Maximum numeric value.
    pub(crate) max_value: Option<f64>,
}

impl PivotCacheField {
    /// Create a new cache field.
    pub fn new(name: impl Into<String>) -> PivotCacheField {
        PivotCacheField {
            name: name.into(),
            num_fmt_id: 0,
            shared_items: Vec::new(),
            contains_string: false,
            contains_number: false,
            contains_integer: false,
            contains_blank: false,
            contains_mixed_types: false,
            min_value: None,
            max_value: None,
        }
    }

    /// Add a string value to shared items.
    pub fn add_string(mut self, value: impl Into<String>) -> PivotCacheField {
        self.shared_items
            .push(PivotCacheValue::String(value.into()));
        self.contains_string = true;
        self
    }

    /// Add a numeric value to shared items.
    pub fn add_number(mut self, value: f64) -> PivotCacheField {
        // Update min/max
        match self.min_value {
            Some(min) if value < min => self.min_value = Some(value),
            None => self.min_value = Some(value),
            _ => {}
        }
        match self.max_value {
            Some(max) if value > max => self.max_value = Some(value),
            None => self.max_value = Some(value),
            _ => {}
        }

        self.shared_items.push(PivotCacheValue::Number(value));
        self.contains_number = true;

        // Check if integer
        if value.fract() == 0.0 {
            self.contains_integer = true;
        }

        self
    }

    /// Add a boolean value to shared items.
    pub fn add_boolean(mut self, value: bool) -> PivotCacheField {
        self.shared_items.push(PivotCacheValue::Boolean(value));
        self
    }

    /// Add a date/time value to shared items.
    pub fn add_datetime(mut self, value: impl Into<String>) -> PivotCacheField {
        self.shared_items
            .push(PivotCacheValue::DateTime(value.into()));
        self
    }

    /// Add a missing/null value to shared items.
    pub fn add_missing(mut self) -> PivotCacheField {
        self.shared_items.push(PivotCacheValue::Missing);
        self.contains_blank = true;
        self
    }

    /// Set the number format ID.
    pub fn set_num_fmt_id(mut self, num_fmt_id: u32) -> PivotCacheField {
        self.num_fmt_id = num_fmt_id;
        self
    }
}

// -----------------------------------------------------------------------
// PivotCacheRecord
// -----------------------------------------------------------------------

/// Represents a record (row) in the pivot cache.
#[derive(Debug, Clone)]
pub struct PivotCacheRecord {
    /// Values in this record (one per field).
    pub(crate) values: Vec<PivotCacheRecordValue>,
}

/// A value in a pivot cache record.
#[derive(Debug, Clone)]
pub enum PivotCacheRecordValue {
    /// Index into shared items.
    SharedItemIndex(u32),

    /// Inline string value.
    String(String),

    /// Inline numeric value.
    Number(f64),

    /// Inline boolean value.
    Boolean(bool),

    /// Inline date value.
    DateTime(String),

    /// Inline error value.
    Error(String),

    /// Missing/null value.
    Missing,
}

impl PivotCacheRecord {
    /// Create a new cache record.
    pub fn new() -> PivotCacheRecord {
        PivotCacheRecord { values: Vec::new() }
    }

    /// Add a shared item index value.
    pub fn add_shared_index(mut self, index: u32) -> PivotCacheRecord {
        self.values
            .push(PivotCacheRecordValue::SharedItemIndex(index));
        self
    }

    /// Add an inline string value.
    pub fn add_string(mut self, value: impl Into<String>) -> PivotCacheRecord {
        self.values
            .push(PivotCacheRecordValue::String(value.into()));
        self
    }

    /// Add an inline numeric value.
    pub fn add_number(mut self, value: f64) -> PivotCacheRecord {
        self.values.push(PivotCacheRecordValue::Number(value));
        self
    }

    /// Add an inline boolean value.
    pub fn add_boolean(mut self, value: bool) -> PivotCacheRecord {
        self.values.push(PivotCacheRecordValue::Boolean(value));
        self
    }

    /// Add an inline date value.
    pub fn add_datetime(mut self, value: impl Into<String>) -> PivotCacheRecord {
        self.values
            .push(PivotCacheRecordValue::DateTime(value.into()));
        self
    }

    /// Add a missing/null value.
    pub fn add_missing(mut self) -> PivotCacheRecord {
        self.values.push(PivotCacheRecordValue::Missing);
        self
    }
}

impl Default for PivotCacheRecord {
    fn default() -> Self {
        Self::new()
    }
}

// -----------------------------------------------------------------------
// PivotCacheDefinition
// -----------------------------------------------------------------------

/// Represents a pivot cache definition.
///
/// The cache definition contains metadata about the pivot cache,
/// including field definitions and source information.
#[derive(Debug, Clone)]
pub struct PivotCacheDefinition {
    pub(crate) writer: Cursor<Vec<u8>>,

    /// Cache ID.
    pub(crate) id: u32,

    /// Source type.
    pub(crate) source_type: PivotCacheSourceType,

    /// Source worksheet name.
    pub(crate) source_sheet: String,

    /// Source range reference.
    pub(crate) source_ref: String,

    /// Source table name (if applicable).
    pub(crate) source_table: Option<String>,

    /// Cache fields.
    pub(crate) fields: Vec<PivotCacheField>,

    /// Record count.
    pub(crate) record_count: u32,

    /// Whether to save data with cache.
    pub(crate) save_data: bool,

    /// Whether to refresh on load.
    pub(crate) refresh_on_load: bool,

    /// Version info.
    pub(crate) created_version: u8,
    pub(crate) refreshed_version: u8,
    pub(crate) min_refreshable_version: u8,

    /// Relationship ID to records file.
    pub(crate) records_rel_id: String,
}

impl PivotCacheDefinition {
    /// Create a new pivot cache definition.
    pub fn new() -> PivotCacheDefinition {
        PivotCacheDefinition {
            writer: Cursor::new(Vec::with_capacity(8192)),
            id: 0,
            source_type: PivotCacheSourceType::Worksheet,
            source_sheet: String::new(),
            source_ref: String::new(),
            source_table: None,
            fields: Vec::new(),
            record_count: 0,
            save_data: true,
            refresh_on_load: false,
            created_version: 6,
            refreshed_version: 6,
            min_refreshable_version: 3,
            records_rel_id: "rId1".to_string(),
        }
    }

    /// Set the source worksheet and range.
    pub fn set_source_range(
        mut self,
        sheet: impl Into<String>,
        range: impl Into<String>,
    ) -> PivotCacheDefinition {
        self.source_sheet = sheet.into();
        self.source_ref = range.into();
        self.source_type = PivotCacheSourceType::Worksheet;
        self
    }

    /// Set the source table name.
    pub fn set_source_table(mut self, table_name: impl Into<String>) -> PivotCacheDefinition {
        self.source_table = Some(table_name.into());
        self
    }

    /// Add a cache field.
    pub fn add_field(mut self, field: PivotCacheField) -> PivotCacheDefinition {
        self.fields.push(field);
        self
    }

    /// Set the record count.
    pub fn set_record_count(mut self, count: u32) -> PivotCacheDefinition {
        self.record_count = count;
        self
    }

    /// Set whether to save data with cache.
    pub fn set_save_data(mut self, save: bool) -> PivotCacheDefinition {
        self.save_data = save;
        self
    }

    /// Set whether to refresh on load.
    pub fn set_refresh_on_load(mut self, refresh: bool) -> PivotCacheDefinition {
        self.refresh_on_load = refresh;
        self
    }

    /// Assemble the XML file for the cache definition.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        self.write_pivot_cache_definition();

        xml_end_tag(&mut self.writer, "pivotCacheDefinition");
    }

    /// Write the <pivotCacheDefinition> element.
    fn write_pivot_cache_definition(&mut self) {
        let record_count = self.record_count.to_string();
        let created_version = self.created_version.to_string();
        let refreshed_version = self.refreshed_version.to_string();
        let min_refreshable_version = self.min_refreshable_version.to_string();

        let mut attributes: Vec<(&str, &str)> = vec![
            (
                "xmlns",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            ),
            (
                "xmlns:r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            ),
            ("r:id", &self.records_rel_id),
            ("recordCount", &record_count),
            ("createdVersion", &created_version),
            ("refreshedVersion", &refreshed_version),
            ("minRefreshableVersion", &min_refreshable_version),
        ];

        if !self.save_data {
            attributes.push(("saveData", "0"));
        }

        if self.refresh_on_load {
            attributes.push(("refreshOnLoad", "1"));
        }

        xml_start_tag(&mut self.writer, "pivotCacheDefinition", &attributes);

        // Write cache source
        self.write_cache_source();

        // Write cache fields
        self.write_cache_fields();
    }

    /// Write the <cacheSource> element.
    fn write_cache_source(&mut self) {
        let type_str = self.source_type.to_type_str();
        let attributes = [("type", type_str)];

        xml_start_tag(&mut self.writer, "cacheSource", &attributes);

        // Write worksheet source
        self.write_worksheet_source();

        xml_end_tag(&mut self.writer, "cacheSource");
    }

    /// Write the <worksheetSource> element.
    fn write_worksheet_source(&mut self) {
        let mut attributes: Vec<(&str, &str)> = Vec::new();

        if let Some(ref table) = self.source_table {
            attributes.push(("name", table.as_str()));
        }

        attributes.push(("sheet", &self.source_sheet));
        attributes.push(("ref", &self.source_ref));

        xml_empty_tag(&mut self.writer, "worksheetSource", &attributes);
    }

    /// Write the <cacheFields> element.
    fn write_cache_fields(&mut self) {
        let count = self.fields.len().to_string();
        let attributes = [("count", count.as_str())];

        xml_start_tag(&mut self.writer, "cacheFields", &attributes);

        for field in self.fields.clone() {
            self.write_cache_field(&field);
        }

        xml_end_tag(&mut self.writer, "cacheFields");
    }

    /// Write a <cacheField> element.
    fn write_cache_field(&mut self, field: &PivotCacheField) {
        let num_fmt_id = field.num_fmt_id.to_string();
        let attributes = [
            ("name", field.name.as_str()),
            ("numFmtId", num_fmt_id.as_str()),
        ];

        xml_start_tag(&mut self.writer, "cacheField", &attributes);

        // Write shared items
        self.write_shared_items(field);

        xml_end_tag(&mut self.writer, "cacheField");
    }

    /// Write the <sharedItems> element.
    fn write_shared_items(&mut self, field: &PivotCacheField) {
        let count = field.shared_items.len().to_string();

        let mut attributes: Vec<(&str, String)> = vec![("count", count)];

        if field.contains_string {
            // String is default, no attribute needed
        }

        if field.contains_number {
            attributes.push(("containsNumber", "1".to_string()));

            if let Some(min) = field.min_value {
                attributes.push(("minValue", min.to_string()));
            }
            if let Some(max) = field.max_value {
                attributes.push(("maxValue", max.to_string()));
            }
        }

        if field.contains_integer {
            attributes.push(("containsInteger", "1".to_string()));
        }

        if field.contains_blank {
            attributes.push(("containsBlank", "1".to_string()));
        }

        if field.contains_mixed_types {
            attributes.push(("containsSemiMixedTypes", "1".to_string()));
        }

        let attr_refs: Vec<(&str, &str)> =
            attributes.iter().map(|(k, v)| (*k, v.as_str())).collect();

        if field.shared_items.is_empty() {
            xml_empty_tag(&mut self.writer, "sharedItems", &attr_refs);
        } else {
            xml_start_tag(&mut self.writer, "sharedItems", &attr_refs);

            for item in &field.shared_items {
                self.write_shared_item(item);
            }

            xml_end_tag(&mut self.writer, "sharedItems");
        }
    }

    /// Write a shared item element.
    fn write_shared_item(&mut self, item: &PivotCacheValue) {
        match item {
            PivotCacheValue::String(s) => {
                xml_empty_tag(&mut self.writer, "s", &[("v", s.as_str())]);
            }
            PivotCacheValue::Number(n) => {
                let v = n.to_string();
                xml_empty_tag(&mut self.writer, "n", &[("v", v.as_str())]);
            }
            PivotCacheValue::Boolean(b) => {
                let v = if *b { "1" } else { "0" };
                xml_empty_tag(&mut self.writer, "b", &[("v", v)]);
            }
            PivotCacheValue::DateTime(d) => {
                xml_empty_tag(&mut self.writer, "d", &[("v", d.as_str())]);
            }
            PivotCacheValue::Error(e) => {
                xml_empty_tag(&mut self.writer, "e", &[("v", e.as_str())]);
            }
            PivotCacheValue::Missing => {
                xml_empty_tag(&mut self.writer, "m", &[] as &[(&str, &str)]);
            }
        }
    }
}

impl Default for PivotCacheDefinition {
    fn default() -> Self {
        Self::new()
    }
}

// -----------------------------------------------------------------------
// PivotCacheRecords
// -----------------------------------------------------------------------

/// Represents the records (data) in a pivot cache.
///
/// This contains the actual data values for the pivot table.
#[derive(Debug, Clone)]
pub struct PivotCacheRecords {
    pub(crate) writer: Cursor<Vec<u8>>,

    /// Records in the cache.
    pub(crate) records: Vec<PivotCacheRecord>,
}

impl PivotCacheRecords {
    /// Create a new pivot cache records collection.
    pub fn new() -> PivotCacheRecords {
        PivotCacheRecords {
            writer: Cursor::new(Vec::with_capacity(16384)),
            records: Vec::new(),
        }
    }

    /// Add a record to the collection.
    pub fn add_record(mut self, record: PivotCacheRecord) -> PivotCacheRecords {
        self.records.push(record);
        self
    }

    /// Get the record count.
    pub fn count(&self) -> u32 {
        self.records.len() as u32
    }

    /// Assemble the XML file for the cache records.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        self.write_pivot_cache_records();

        xml_end_tag(&mut self.writer, "pivotCacheRecords");
    }

    /// Write the <pivotCacheRecords> element.
    fn write_pivot_cache_records(&mut self) {
        let count = self.records.len().to_string();

        let attributes = [
            (
                "xmlns",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            ),
            ("count", count.as_str()),
        ];

        xml_start_tag(&mut self.writer, "pivotCacheRecords", &attributes);

        for record in self.records.clone() {
            self.write_record(&record);
        }
    }

    /// Write an <r> (record) element.
    fn write_record(&mut self, record: &PivotCacheRecord) {
        xml_start_tag(&mut self.writer, "r", &[] as &[(&str, &str)]);

        for value in &record.values {
            self.write_record_value(value);
        }

        xml_end_tag(&mut self.writer, "r");
    }

    /// Write a record value element.
    fn write_record_value(&mut self, value: &PivotCacheRecordValue) {
        match value {
            PivotCacheRecordValue::SharedItemIndex(i) => {
                let v = i.to_string();
                xml_empty_tag(&mut self.writer, "x", &[("v", v.as_str())]);
            }
            PivotCacheRecordValue::String(s) => {
                xml_empty_tag(&mut self.writer, "s", &[("v", s.as_str())]);
            }
            PivotCacheRecordValue::Number(n) => {
                let v = n.to_string();
                xml_empty_tag(&mut self.writer, "n", &[("v", v.as_str())]);
            }
            PivotCacheRecordValue::Boolean(b) => {
                let v = if *b { "1" } else { "0" };
                xml_empty_tag(&mut self.writer, "b", &[("v", v)]);
            }
            PivotCacheRecordValue::DateTime(d) => {
                xml_empty_tag(&mut self.writer, "d", &[("v", d.as_str())]);
            }
            PivotCacheRecordValue::Error(e) => {
                xml_empty_tag(&mut self.writer, "e", &[("v", e.as_str())]);
            }
            PivotCacheRecordValue::Missing => {
                xml_empty_tag(&mut self.writer, "m", &[] as &[(&str, &str)]);
            }
        }
    }
}

impl Default for PivotCacheRecords {
    fn default() -> Self {
        Self::new()
    }
}
