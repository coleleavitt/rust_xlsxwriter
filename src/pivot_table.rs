// pivot_table - A module for creating Excel PivotTables.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]
#![allow(dead_code)] // Infrastructure for future pivot table API

mod tests;

use std::io::Cursor;

use crate::xmlwriter::{xml_declaration, xml_empty_tag, xml_end_tag, xml_start_tag};
#[allow(unused_imports)]
use crate::{CellRange, ColNum, RowNum, XlsxError};

// -----------------------------------------------------------------------
// Enums
// -----------------------------------------------------------------------

/// The axis position for a pivot table field.
///
/// Determines where a field is placed in the pivot table layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PivotFieldAxis {
    /// Field is not placed on any axis.
    #[default]
    None,

    /// Field is placed on the row axis.
    Row,

    /// Field is placed on the column axis.
    Column,

    /// Field is placed on the page/filter axis.
    Page,

    /// Field is used as a data/values field.
    Values,
}

impl PivotFieldAxis {
    /// Convert to XML axis attribute value.
    pub(crate) fn to_axis_str(&self) -> Option<&'static str> {
        match self {
            PivotFieldAxis::None => None,
            PivotFieldAxis::Row => Some("axisRow"),
            PivotFieldAxis::Column => Some("axisCol"),
            PivotFieldAxis::Page => Some("axisPage"),
            PivotFieldAxis::Values => Some("axisValues"),
        }
    }
}

/// The aggregation function for a pivot table data field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DataFieldFunction {
    /// Sum of values.
    #[default]
    Sum,

    /// Count of values.
    Count,

    /// Average of values.
    Average,

    /// Maximum value.
    Max,

    /// Minimum value.
    Min,

    /// Product of values.
    Product,

    /// Count of numeric values.
    CountNums,

    /// Sample standard deviation.
    StdDev,

    /// Population standard deviation.
    StdDevP,

    /// Sample variance.
    Var,

    /// Population variance.
    VarP,
}

impl DataFieldFunction {
    /// Convert to XML subtotal attribute value.
    pub(crate) fn to_subtotal_str(&self) -> &'static str {
        match self {
            DataFieldFunction::Sum => "sum",
            DataFieldFunction::Count => "count",
            DataFieldFunction::Average => "average",
            DataFieldFunction::Max => "max",
            DataFieldFunction::Min => "min",
            DataFieldFunction::Product => "product",
            DataFieldFunction::CountNums => "countNums",
            DataFieldFunction::StdDev => "stdDev",
            DataFieldFunction::StdDevP => "stdDevp",
            DataFieldFunction::Var => "var",
            DataFieldFunction::VarP => "varp",
        }
    }
}

/// The "Show Data As" calculation type for a pivot table data field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ShowDataAs {
    /// Normal values (no calculation).
    #[default]
    Normal,

    /// Difference from base value.
    Difference,

    /// Percentage difference from base.
    PercentDifference,

    /// Percentage of row total.
    PercentOfRow,

    /// Percentage of column total.
    PercentOfColumn,

    /// Percentage of grand total.
    PercentOfTotal,

    /// Index calculation.
    Index,

    /// Percentage of parent row total.
    PercentOfParentRow,

    /// Percentage of parent column total.
    PercentOfParentColumn,

    /// Percentage of parent total.
    PercentOfParent,

    /// Running total.
    RunningTotal,

    /// Percentage running total.
    PercentRunningTotal,

    /// Rank ascending.
    RankAscending,

    /// Rank descending.
    RankDescending,
}

impl ShowDataAs {
    /// Convert to XML showDataAs attribute value.
    pub(crate) fn to_show_data_as_str(&self) -> Option<&'static str> {
        match self {
            ShowDataAs::Normal => None,
            ShowDataAs::Difference => Some("difference"),
            ShowDataAs::PercentDifference => Some("percentDiff"),
            ShowDataAs::PercentOfRow => Some("percentOfRow"),
            ShowDataAs::PercentOfColumn => Some("percentOfCol"),
            ShowDataAs::PercentOfTotal => Some("percentOfTotal"),
            ShowDataAs::Index => Some("index"),
            ShowDataAs::PercentOfParentRow => Some("percentOfParentRow"),
            ShowDataAs::PercentOfParentColumn => Some("percentOfParentCol"),
            ShowDataAs::PercentOfParent => Some("percentOfParent"),
            ShowDataAs::RunningTotal => Some("runTotal"),
            ShowDataAs::PercentRunningTotal => Some("percentOfRunningTotal"),
            ShowDataAs::RankAscending => Some("rankAscending"),
            ShowDataAs::RankDescending => Some("rankDescending"),
        }
    }
}

/// The grouping type for date fields in a pivot table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DateGroupBy {
    /// Group by years.
    Years,

    /// Group by quarters.
    Quarters,

    /// Group by months.
    Months,

    /// Group by days.
    Days,

    /// Group by hours.
    Hours,

    /// Group by minutes.
    Minutes,

    /// Group by seconds.
    Seconds,
}

// -----------------------------------------------------------------------
// PivotField
// -----------------------------------------------------------------------

/// Represents a field in a pivot table.
///
/// A field corresponds to a column in the source data and can be placed
/// on row, column, page, or values axes.
#[derive(Debug, Clone)]
pub struct PivotField {
    /// Field name (from source column header).
    pub(crate) name: String,

    /// Field index in the cache.
    pub(crate) index: u32,

    /// Axis placement.
    pub(crate) axis: PivotFieldAxis,

    /// Whether this is a data field.
    pub(crate) data_field: bool,

    /// Show all items.
    pub(crate) show_all: bool,

    /// Compact display mode.
    pub(crate) compact: bool,

    /// Outline display mode.
    pub(crate) outline: bool,

    /// Subtotal at top.
    pub(crate) subtotal_top: bool,

    /// Insert blank row after items.
    pub(crate) insert_blank_row: bool,

    /// Items in this field.
    pub(crate) items: Vec<PivotFieldItem>,
}

impl PivotField {
    /// Create a new pivot field.
    pub fn new(name: impl Into<String>) -> PivotField {
        PivotField {
            name: name.into(),
            index: 0,
            axis: PivotFieldAxis::None,
            data_field: false,
            show_all: false,
            compact: true,
            outline: true,
            subtotal_top: true,
            insert_blank_row: false,
            items: Vec::new(),
        }
    }

    /// Set the axis for this field.
    pub fn set_axis(mut self, axis: PivotFieldAxis) -> PivotField {
        self.axis = axis;
        self
    }

    /// Set whether this is a data field.
    pub fn set_data_field(mut self, is_data_field: bool) -> PivotField {
        self.data_field = is_data_field;
        self
    }

    /// Set compact display mode.
    pub fn set_compact(mut self, compact: bool) -> PivotField {
        self.compact = compact;
        self
    }

    /// Set outline display mode.
    pub fn set_outline(mut self, outline: bool) -> PivotField {
        self.outline = outline;
        self
    }

    /// Add an item to this field.
    pub fn add_item(mut self, item: PivotFieldItem) -> PivotField {
        self.items.push(item);
        self
    }
}

/// Represents an item (value) within a pivot field.
#[derive(Debug, Clone)]
pub struct PivotFieldItem {
    /// Item index in shared items.
    pub(crate) index: u32,

    /// Whether the item is hidden.
    pub(crate) hidden: bool,

    /// Item type (data, default, grand, etc.).
    pub(crate) item_type: Option<String>,
}

impl PivotFieldItem {
    /// Create a new field item.
    pub fn new(index: u32) -> PivotFieldItem {
        PivotFieldItem {
            index,
            hidden: false,
            item_type: None,
        }
    }

    /// Set whether the item is hidden.
    pub fn set_hidden(mut self, hidden: bool) -> PivotFieldItem {
        self.hidden = hidden;
        self
    }

    /// Set the item type.
    pub fn set_type(mut self, item_type: impl Into<String>) -> PivotFieldItem {
        self.item_type = Some(item_type.into());
        self
    }
}

// -----------------------------------------------------------------------
// PivotDataField
// -----------------------------------------------------------------------

/// Represents a data (values) field in a pivot table.
///
/// Data fields display aggregated values in the pivot table body.
#[derive(Debug, Clone)]
pub struct PivotDataField {
    /// Display name for this data field.
    pub(crate) name: String,

    /// Index of the source field.
    pub(crate) field_index: u32,

    /// Aggregation function.
    pub(crate) function: DataFieldFunction,

    /// Show Data As calculation.
    pub(crate) show_data_as: ShowDataAs,

    /// Base field for calculations.
    pub(crate) base_field: u32,

    /// Base item for calculations.
    pub(crate) base_item: u32,

    /// Number format ID.
    pub(crate) num_fmt_id: u32,
}

impl PivotDataField {
    /// Create a new data field.
    ///
    /// # Arguments
    ///
    /// * `name` - Display name for the data field (e.g., "Sum of Sales").
    /// * `field_index` - Index of the source field in the cache.
    pub fn new(name: impl Into<String>, field_index: u32) -> PivotDataField {
        PivotDataField {
            name: name.into(),
            field_index,
            function: DataFieldFunction::Sum,
            show_data_as: ShowDataAs::Normal,
            base_field: 0,
            base_item: 0,
            num_fmt_id: 0,
        }
    }

    /// Set the aggregation function.
    pub fn set_function(mut self, function: DataFieldFunction) -> PivotDataField {
        self.function = function;
        self
    }

    /// Set the Show Data As calculation.
    pub fn set_show_data_as(mut self, show_data_as: ShowDataAs) -> PivotDataField {
        self.show_data_as = show_data_as;
        self
    }

    /// Set the base field for calculations.
    pub fn set_base_field(mut self, base_field: u32) -> PivotDataField {
        self.base_field = base_field;
        self
    }

    /// Set the base item for calculations.
    pub fn set_base_item(mut self, base_item: u32) -> PivotDataField {
        self.base_item = base_item;
        self
    }

    /// Set the number format ID.
    pub fn set_num_fmt_id(mut self, num_fmt_id: u32) -> PivotDataField {
        self.num_fmt_id = num_fmt_id;
        self
    }
}

// -----------------------------------------------------------------------
// PivotTable
// -----------------------------------------------------------------------

/// Represents an Excel PivotTable.
///
/// A pivot table summarizes data from a source range or table, allowing
/// users to analyze and explore data interactively.
///
/// # Examples
///
/// ```ignore
/// use rust_xlsxwriter::{PivotTable, PivotField, PivotDataField, DataFieldFunction, Workbook};
///
/// let mut workbook = Workbook::new();
/// let worksheet = workbook.add_worksheet();
///
/// // Add source data
/// worksheet.write(0, 0, "Region")?;
/// worksheet.write(0, 1, "Sales")?;
/// worksheet.write(1, 0, "North")?;
/// worksheet.write(1, 1, 1000)?;
/// worksheet.write(2, 0, "South")?;
/// worksheet.write(2, 1, 2000)?;
///
/// // Create pivot table
/// let pivot = PivotTable::new()
///     .set_name("SalesPivot")
///     .set_data_range("Sheet1!A1:B3")
///     .add_row_field(PivotField::new("Region"))
///     .add_data_field(PivotDataField::new("Sum of Sales", 1).set_function(DataFieldFunction::Sum));
///
/// worksheet.add_pivot_table(5, 0, &pivot)?;
/// ```
#[derive(Debug, Clone)]
pub struct PivotTable {
    pub(crate) writer: Cursor<Vec<u8>>,

    /// Pivot table name.
    pub(crate) name: String,

    /// Pivot table ID.
    pub(crate) id: u32,

    /// Cache ID reference.
    pub(crate) cache_id: u32,

    /// Location cell reference (top-left corner).
    pub(crate) location_ref: String,

    /// Data range reference.
    pub(crate) data_range: String,

    /// Row fields.
    pub(crate) row_fields: Vec<PivotField>,

    /// Column fields.
    pub(crate) column_fields: Vec<PivotField>,

    /// Page (filter) fields.
    pub(crate) page_fields: Vec<PivotField>,

    /// Data (values) fields.
    pub(crate) data_fields: Vec<PivotDataField>,

    /// All fields in the pivot table.
    pub(crate) fields: Vec<PivotField>,

    /// Show row grand totals.
    pub(crate) row_grand_totals: bool,

    /// Show column grand totals.
    pub(crate) col_grand_totals: bool,

    /// Compact layout mode.
    pub(crate) compact: bool,

    /// Outline layout mode.
    pub(crate) outline: bool,

    /// Data caption text.
    pub(crate) data_caption: String,

    /// Version info.
    pub(crate) created_version: u8,
    pub(crate) updated_version: u8,
    pub(crate) min_refreshable_version: u8,
}

impl PivotTable {
    /// Create a new pivot table.
    pub fn new() -> PivotTable {
        PivotTable {
            writer: Cursor::new(Vec::with_capacity(8192)),
            name: String::new(),
            id: 0,
            cache_id: 0,
            location_ref: String::new(),
            data_range: String::new(),
            row_fields: Vec::new(),
            column_fields: Vec::new(),
            page_fields: Vec::new(),
            data_fields: Vec::new(),
            fields: Vec::new(),
            row_grand_totals: true,
            col_grand_totals: true,
            compact: true,
            outline: false,
            data_caption: "Data".to_string(),
            created_version: 6,
            updated_version: 6,
            min_refreshable_version: 3,
        }
    }

    /// Set the pivot table name.
    pub fn set_name(mut self, name: impl Into<String>) -> PivotTable {
        self.name = name.into();
        self
    }

    /// Set the data source range.
    ///
    /// The range should be in the format "Sheet1!A1:D100" or a table name.
    pub fn set_data_range(mut self, range: impl Into<String>) -> PivotTable {
        self.data_range = range.into();
        self
    }

    /// Add a row field to the pivot table.
    pub fn add_row_field(mut self, field: PivotField) -> PivotTable {
        let mut field = field;
        field.axis = PivotFieldAxis::Row;
        self.row_fields.push(field);
        self
    }

    /// Add a column field to the pivot table.
    pub fn add_column_field(mut self, field: PivotField) -> PivotTable {
        let mut field = field;
        field.axis = PivotFieldAxis::Column;
        self.column_fields.push(field);
        self
    }

    /// Add a page (filter) field to the pivot table.
    pub fn add_page_field(mut self, field: PivotField) -> PivotTable {
        let mut field = field;
        field.axis = PivotFieldAxis::Page;
        self.page_fields.push(field);
        self
    }

    /// Add a data (values) field to the pivot table.
    pub fn add_data_field(mut self, data_field: PivotDataField) -> PivotTable {
        self.data_fields.push(data_field);
        self
    }

    /// Set whether to show row grand totals.
    pub fn set_row_grand_totals(mut self, show: bool) -> PivotTable {
        self.row_grand_totals = show;
        self
    }

    /// Set whether to show column grand totals.
    pub fn set_col_grand_totals(mut self, show: bool) -> PivotTable {
        self.col_grand_totals = show;
        self
    }

    /// Set the data caption text.
    pub fn set_data_caption(mut self, caption: impl Into<String>) -> PivotTable {
        self.data_caption = caption.into();
        self
    }

    /// Set compact layout mode.
    pub fn set_compact(mut self, compact: bool) -> PivotTable {
        self.compact = compact;
        self
    }

    /// Set outline layout mode.
    pub fn set_outline(mut self, outline: bool) -> PivotTable {
        self.outline = outline;
        self
    }

    /// Assemble the XML file for the pivot table definition.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        self.write_pivot_table_definition();

        xml_end_tag(&mut self.writer, "pivotTableDefinition");
    }

    /// Write the <pivotTableDefinition> element.
    fn write_pivot_table_definition(&mut self) {
        let mut attributes: Vec<(&str, String)> = vec![
            (
                "xmlns",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string(),
            ),
            ("name", self.name.clone()),
            ("cacheId", self.cache_id.to_string()),
            ("dataCaption", self.data_caption.clone()),
            ("updatedVersion", self.updated_version.to_string()),
            (
                "minRefreshableVersion",
                self.min_refreshable_version.to_string(),
            ),
            ("createdVersion", self.created_version.to_string()),
        ];

        if !self.row_grand_totals {
            attributes.push(("rowGrandTotals", "0".to_string()));
        }
        if !self.col_grand_totals {
            attributes.push(("colGrandTotals", "0".to_string()));
        }
        if self.outline {
            attributes.push(("outline", "1".to_string()));
        }
        if !self.compact {
            attributes.push(("compact", "0".to_string()));
        }

        let attr_refs: Vec<(&str, &str)> =
            attributes.iter().map(|(k, v)| (*k, v.as_str())).collect();

        xml_start_tag(&mut self.writer, "pivotTableDefinition", &attr_refs);

        // Write location
        self.write_location();

        // Write pivot fields
        self.write_pivot_fields();

        // Write row fields
        if !self.row_fields.is_empty() {
            self.write_row_fields();
        }

        // Write column fields
        if !self.column_fields.is_empty() {
            self.write_col_fields();
        }

        // Write page fields
        if !self.page_fields.is_empty() {
            self.write_page_fields();
        }

        // Write data fields
        if !self.data_fields.is_empty() {
            self.write_data_fields();
        }

        // Write pivot table style info
        self.write_pivot_table_style_info();
    }

    /// Write the <location> element.
    fn write_location(&mut self) {
        let attributes = [
            ("ref", self.location_ref.as_str()),
            ("firstHeaderRow", "1"),
            ("firstDataRow", "1"),
            ("firstDataCol", "0"),
        ];

        xml_empty_tag(&mut self.writer, "location", &attributes);
    }

    /// Write the <pivotFields> element.
    fn write_pivot_fields(&mut self) {
        let count = self.fields.len().to_string();
        let attributes = [("count", count.as_str())];

        xml_start_tag(&mut self.writer, "pivotFields", &attributes);

        for field in self.fields.clone() {
            self.write_pivot_field(&field);
        }

        xml_end_tag(&mut self.writer, "pivotFields");
    }

    /// Write a <pivotField> element.
    fn write_pivot_field(&mut self, field: &PivotField) {
        let mut attributes: Vec<(&str, String)> = Vec::new();

        if let Some(axis) = field.axis.to_axis_str() {
            attributes.push(("axis", axis.to_string()));
        }

        if field.data_field {
            attributes.push(("dataField", "1".to_string()));
        }

        attributes.push((
            "showAll",
            if field.show_all { "1" } else { "0" }.to_string(),
        ));

        if !field.compact {
            attributes.push(("compact", "0".to_string()));
        }

        if field.outline {
            attributes.push(("outline", "1".to_string()));
        }

        if field.subtotal_top {
            attributes.push(("subtotalTop", "1".to_string()));
        }

        let attr_refs: Vec<(&str, &str)> =
            attributes.iter().map(|(k, v)| (*k, v.as_str())).collect();

        if field.items.is_empty() {
            xml_empty_tag(&mut self.writer, "pivotField", &attr_refs);
        } else {
            xml_start_tag(&mut self.writer, "pivotField", &attr_refs);
            self.write_items(&field.items);
            xml_end_tag(&mut self.writer, "pivotField");
        }
    }

    /// Write the <items> element.
    fn write_items(&mut self, items: &[PivotFieldItem]) {
        let count = items.len().to_string();
        let attributes = [("count", count.as_str())];

        xml_start_tag(&mut self.writer, "items", &attributes);

        for item in items {
            self.write_item(item);
        }

        xml_end_tag(&mut self.writer, "items");
    }

    /// Write an <item> element.
    fn write_item(&mut self, item: &PivotFieldItem) {
        let mut attributes: Vec<(&str, String)> = Vec::new();

        if let Some(ref t) = item.item_type {
            attributes.push(("t", t.clone()));
        } else {
            attributes.push(("x", item.index.to_string()));
        }

        if item.hidden {
            attributes.push(("h", "1".to_string()));
        }

        let attr_refs: Vec<(&str, &str)> =
            attributes.iter().map(|(k, v)| (*k, v.as_str())).collect();

        xml_empty_tag(&mut self.writer, "item", &attr_refs);
    }

    /// Write the <rowFields> element.
    fn write_row_fields(&mut self) {
        let count = self.row_fields.len().to_string();
        let attributes = [("count", count.as_str())];

        xml_start_tag(&mut self.writer, "rowFields", &attributes);

        for field in &self.row_fields {
            let x = field.index.to_string();
            xml_empty_tag(&mut self.writer, "field", &[("x", x.as_str())]);
        }

        xml_end_tag(&mut self.writer, "rowFields");
    }

    /// Write the <colFields> element.
    fn write_col_fields(&mut self) {
        let count = self.column_fields.len().to_string();
        let attributes = [("count", count.as_str())];

        xml_start_tag(&mut self.writer, "colFields", &attributes);

        for field in &self.column_fields {
            let x = field.index.to_string();
            xml_empty_tag(&mut self.writer, "field", &[("x", x.as_str())]);
        }

        xml_end_tag(&mut self.writer, "colFields");
    }

    /// Write the <pageFields> element.
    fn write_page_fields(&mut self) {
        let count = self.page_fields.len().to_string();
        let attributes = [("count", count.as_str())];

        xml_start_tag(&mut self.writer, "pageFields", &attributes);

        for field in &self.page_fields {
            let fld = field.index.to_string();
            xml_empty_tag(&mut self.writer, "pageField", &[("fld", fld.as_str())]);
        }

        xml_end_tag(&mut self.writer, "pageFields");
    }

    /// Write the <dataFields> element.
    fn write_data_fields(&mut self) {
        let count = self.data_fields.len().to_string();
        let attributes = [("count", count.as_str())];

        xml_start_tag(&mut self.writer, "dataFields", &attributes);

        for data_field in self.data_fields.clone() {
            self.write_data_field(&data_field);
        }

        xml_end_tag(&mut self.writer, "dataFields");
    }

    /// Write a <dataField> element.
    fn write_data_field(&mut self, data_field: &PivotDataField) {
        let fld = data_field.field_index.to_string();
        let base_field = data_field.base_field.to_string();
        let base_item = data_field.base_item.to_string();

        let mut attributes: Vec<(&str, &str)> = vec![
            ("name", data_field.name.as_str()),
            ("fld", fld.as_str()),
            ("subtotal", data_field.function.to_subtotal_str()),
            ("baseField", base_field.as_str()),
            ("baseItem", base_item.as_str()),
        ];

        if let Some(show_data_as) = data_field.show_data_as.to_show_data_as_str() {
            attributes.push(("showDataAs", show_data_as));
        }

        xml_empty_tag(&mut self.writer, "dataField", &attributes);
    }

    /// Write the <pivotTableStyleInfo> element.
    fn write_pivot_table_style_info(&mut self) {
        let attributes = [
            ("name", "PivotStyleLight16"),
            ("showRowHeaders", "1"),
            ("showColHeaders", "1"),
            ("showRowStripes", "0"),
            ("showColStripes", "0"),
            ("showLastColumn", "1"),
        ];

        xml_empty_tag(&mut self.writer, "pivotTableStyleInfo", &attributes);
    }
}

impl Default for PivotTable {
    fn default() -> Self {
        Self::new()
    }
}
