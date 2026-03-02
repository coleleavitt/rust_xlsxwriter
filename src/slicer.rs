// slicer - A module for creating Excel Slicers.
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

/// The source type for a slicer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SlicerSourceType {
    /// Slicer is connected to a table.
    #[default]
    Table,

    /// Slicer is connected to a pivot table.
    PivotTable,
}

/// The cross filter behavior for a slicer.
///
/// Determines how items with no data are displayed in the slicer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SlicerCrossFilter {
    /// No cross filter behavior.
    #[default]
    None,

    /// Show items with data at the top.
    ShowItemsWithDataAtTop,

    /// Show items with no data (dimmed).
    ShowItemsWithNoData,
}

impl SlicerCrossFilter {
    /// Convert to XML crossFilter attribute value.
    pub(crate) fn to_cross_filter_str(&self) -> &'static str {
        match self {
            SlicerCrossFilter::None => "none",
            SlicerCrossFilter::ShowItemsWithDataAtTop => "showItemsWithDataAtTop",
            SlicerCrossFilter::ShowItemsWithNoData => "showItemsWithNoData",
        }
    }
}

/// The sort order for slicer items.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SlicerSortOrder {
    /// Ascending sort order.
    #[default]
    Ascending,

    /// Descending sort order.
    Descending,
}

impl SlicerSortOrder {
    /// Convert to XML sortOrder attribute value.
    pub(crate) fn to_sort_order_str(&self) -> &'static str {
        match self {
            SlicerSortOrder::Ascending => "ascending",
            SlicerSortOrder::Descending => "descending",
        }
    }
}

/// Built-in slicer styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SlicerStyle {
    /// Light style 1.
    SlicerStyleLight1,

    /// Light style 2.
    SlicerStyleLight2,

    /// Light style 3.
    SlicerStyleLight3,

    /// Light style 4.
    #[default]
    SlicerStyleLight4,

    /// Light style 5.
    SlicerStyleLight5,

    /// Light style 6.
    SlicerStyleLight6,

    /// Dark style 1.
    SlicerStyleDark1,

    /// Dark style 2.
    SlicerStyleDark2,

    /// Dark style 3.
    SlicerStyleDark3,

    /// Dark style 4.
    SlicerStyleDark4,

    /// Dark style 5.
    SlicerStyleDark5,

    /// Dark style 6.
    SlicerStyleDark6,

    /// Other style (custom name).
    Other,
}

impl SlicerStyle {
    /// Convert to XML style name.
    pub(crate) fn to_style_str(&self) -> &'static str {
        match self {
            SlicerStyle::SlicerStyleLight1 => "SlicerStyleLight1",
            SlicerStyle::SlicerStyleLight2 => "SlicerStyleLight2",
            SlicerStyle::SlicerStyleLight3 => "SlicerStyleLight3",
            SlicerStyle::SlicerStyleLight4 => "SlicerStyleLight4",
            SlicerStyle::SlicerStyleLight5 => "SlicerStyleLight5",
            SlicerStyle::SlicerStyleLight6 => "SlicerStyleLight6",
            SlicerStyle::SlicerStyleDark1 => "SlicerStyleDark1",
            SlicerStyle::SlicerStyleDark2 => "SlicerStyleDark2",
            SlicerStyle::SlicerStyleDark3 => "SlicerStyleDark3",
            SlicerStyle::SlicerStyleDark4 => "SlicerStyleDark4",
            SlicerStyle::SlicerStyleDark5 => "SlicerStyleDark5",
            SlicerStyle::SlicerStyleDark6 => "SlicerStyleDark6",
            SlicerStyle::Other => "SlicerStyleLight4",
        }
    }
}

// -----------------------------------------------------------------------
// SlicerItem
// -----------------------------------------------------------------------

/// Represents an item in a slicer.
#[derive(Debug, Clone)]
pub struct SlicerItem {
    /// Item index in the source data.
    pub(crate) index: u32,

    /// Whether the item is selected (visible in filter).
    pub(crate) selected: bool,

    /// Whether the item has no data.
    pub(crate) no_data: bool,
}

impl SlicerItem {
    /// Create a new slicer item.
    pub fn new(index: u32) -> SlicerItem {
        SlicerItem {
            index,
            selected: true,
            no_data: false,
        }
    }

    /// Set whether the item is selected.
    pub fn set_selected(mut self, selected: bool) -> SlicerItem {
        self.selected = selected;
        self
    }

    /// Set whether the item has no data.
    pub fn set_no_data(mut self, no_data: bool) -> SlicerItem {
        self.no_data = no_data;
        self
    }
}

// -----------------------------------------------------------------------
// SlicerCache (Base)
// -----------------------------------------------------------------------

/// Base structure for slicer cache.
#[derive(Debug, Clone)]
pub struct SlicerCacheBase {
    /// Cache name (matches slicer name).
    pub(crate) name: String,

    /// Source name (column/field name).
    pub(crate) source_name: String,

    /// Sort order for items.
    pub(crate) sort_order: SlicerSortOrder,

    /// Cross filter behavior.
    pub(crate) cross_filter: SlicerCrossFilter,

    /// Use custom list for sorting.
    pub(crate) custom_list_sort: bool,

    /// Hide items with no data.
    pub(crate) hide_items_with_no_data: bool,
}

impl SlicerCacheBase {
    /// Create a new slicer cache base.
    pub fn new(name: impl Into<String>, source_name: impl Into<String>) -> SlicerCacheBase {
        SlicerCacheBase {
            name: name.into(),
            source_name: source_name.into(),
            sort_order: SlicerSortOrder::Ascending,
            cross_filter: SlicerCrossFilter::None,
            custom_list_sort: true,
            hide_items_with_no_data: false,
        }
    }
}

// -----------------------------------------------------------------------
// TableSlicerCache
// -----------------------------------------------------------------------

/// Represents a slicer cache for table slicers.
#[derive(Debug, Clone)]
pub struct TableSlicerCache {
    pub(crate) writer: Cursor<Vec<u8>>,

    /// Base cache properties.
    pub(crate) base: SlicerCacheBase,

    /// Table ID reference.
    pub(crate) table_id: u32,

    /// Column ID reference.
    pub(crate) column_id: u32,
}

impl TableSlicerCache {
    /// Create a new table slicer cache.
    pub fn new(
        name: impl Into<String>,
        source_name: impl Into<String>,
        table_id: u32,
        column_id: u32,
    ) -> TableSlicerCache {
        TableSlicerCache {
            writer: Cursor::new(Vec::with_capacity(4096)),
            base: SlicerCacheBase::new(name, source_name),
            table_id,
            column_id,
        }
    }

    /// Set the sort order.
    pub fn set_sort_order(mut self, sort_order: SlicerSortOrder) -> TableSlicerCache {
        self.base.sort_order = sort_order;
        self
    }

    /// Set the cross filter behavior.
    pub fn set_cross_filter(mut self, cross_filter: SlicerCrossFilter) -> TableSlicerCache {
        self.base.cross_filter = cross_filter;
        self
    }

    /// Assemble the XML file for the table slicer cache.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        self.write_slicer_cache_definition();

        xml_end_tag(&mut self.writer, "slicerCacheDefinition");
    }

    /// Write the <slicerCacheDefinition> element.
    fn write_slicer_cache_definition(&mut self) {
        let attributes = [
            (
                "xmlns",
                "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
            ),
            (
                "xmlns:mc",
                "http://schemas.openxmlformats.org/markup-compatibility/2006",
            ),
            ("mc:Ignorable", "x"),
            (
                "xmlns:x",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            ),
            ("name", &self.base.name),
            ("sourceName", &self.base.source_name),
        ];

        xml_start_tag(&mut self.writer, "slicerCacheDefinition", &attributes);

        // Write extLst with tableSlicerCache
        self.write_ext_lst();
    }

    /// Write the <extLst> element.
    fn write_ext_lst(&mut self) {
        xml_start_tag(&mut self.writer, "extLst", &[] as &[(&str, &str)]);

        let attributes = [("uri", "{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}")];
        xml_start_tag(&mut self.writer, "x:ext", &attributes);

        self.write_table_slicer_cache();

        xml_end_tag(&mut self.writer, "x:ext");
        xml_end_tag(&mut self.writer, "extLst");
    }

    /// Write the <x15:tableSlicerCache> element.
    fn write_table_slicer_cache(&mut self) {
        let table_id = self.table_id.to_string();
        let column = self.column_id.to_string();

        let attributes = [
            (
                "xmlns:x15",
                "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
            ),
            ("tableId", table_id.as_str()),
            ("column", column.as_str()),
            ("sortOrder", self.base.sort_order.to_sort_order_str()),
            ("crossFilter", self.base.cross_filter.to_cross_filter_str()),
            (
                "customListSort",
                if self.base.custom_list_sort { "1" } else { "0" },
            ),
        ];

        xml_empty_tag(&mut self.writer, "x15:tableSlicerCache", &attributes);
    }
}

// -----------------------------------------------------------------------
// PivotTableSlicerCache
// -----------------------------------------------------------------------

/// Represents a slicer cache for pivot table slicers.
#[derive(Debug, Clone)]
pub struct PivotTableSlicerCache {
    pub(crate) writer: Cursor<Vec<u8>>,

    /// Base cache properties.
    pub(crate) base: SlicerCacheBase,

    /// Pivot cache ID.
    pub(crate) pivot_cache_id: u32,

    /// Pivot tables attached to this cache.
    pub(crate) pivot_tables: Vec<PivotTableSlicerRef>,

    /// Items in the slicer.
    pub(crate) items: Vec<SlicerItem>,
}

/// Reference to a pivot table for slicer.
#[derive(Debug, Clone)]
pub struct PivotTableSlicerRef {
    /// Pivot table name.
    pub(crate) name: String,

    /// Tab (worksheet) ID.
    pub(crate) tab_id: u32,
}

impl PivotTableSlicerRef {
    /// Create a new pivot table reference.
    pub fn new(name: impl Into<String>, tab_id: u32) -> PivotTableSlicerRef {
        PivotTableSlicerRef {
            name: name.into(),
            tab_id,
        }
    }
}

impl PivotTableSlicerCache {
    /// Create a new pivot table slicer cache.
    pub fn new(
        name: impl Into<String>,
        source_name: impl Into<String>,
        pivot_cache_id: u32,
    ) -> PivotTableSlicerCache {
        PivotTableSlicerCache {
            writer: Cursor::new(Vec::with_capacity(4096)),
            base: SlicerCacheBase::new(name, source_name),
            pivot_cache_id,
            pivot_tables: Vec::new(),
            items: Vec::new(),
        }
    }

    /// Add a pivot table reference.
    pub fn add_pivot_table(mut self, pivot_ref: PivotTableSlicerRef) -> PivotTableSlicerCache {
        self.pivot_tables.push(pivot_ref);
        self
    }

    /// Add an item to the slicer.
    pub fn add_item(mut self, item: SlicerItem) -> PivotTableSlicerCache {
        self.items.push(item);
        self
    }

    /// Set the sort order.
    pub fn set_sort_order(mut self, sort_order: SlicerSortOrder) -> PivotTableSlicerCache {
        self.base.sort_order = sort_order;
        self
    }

    /// Set the cross filter behavior.
    pub fn set_cross_filter(mut self, cross_filter: SlicerCrossFilter) -> PivotTableSlicerCache {
        self.base.cross_filter = cross_filter;
        self
    }

    /// Assemble the XML file for the pivot table slicer cache.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        self.write_slicer_cache_definition();

        xml_end_tag(&mut self.writer, "slicerCacheDefinition");
    }

    /// Write the <slicerCacheDefinition> element.
    fn write_slicer_cache_definition(&mut self) {
        let attributes = [
            (
                "xmlns",
                "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
            ),
            (
                "xmlns:mc",
                "http://schemas.openxmlformats.org/markup-compatibility/2006",
            ),
            ("mc:Ignorable", "x"),
            (
                "xmlns:x",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            ),
            (
                "xmlns:x14",
                "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
            ),
            ("name", &self.base.name),
            ("sourceName", &self.base.source_name),
        ];

        xml_start_tag(&mut self.writer, "slicerCacheDefinition", &attributes);

        // Write pivot tables
        if !self.pivot_tables.is_empty() {
            self.write_pivot_tables();
        }

        // Write data (tabular)
        self.write_data();
    }

    /// Write the <x14:pivotTables> element.
    fn write_pivot_tables(&mut self) {
        xml_start_tag(&mut self.writer, "x14:pivotTables", &[] as &[(&str, &str)]);

        for pt_ref in &self.pivot_tables {
            let tab_id = pt_ref.tab_id.to_string();
            let attributes = [("name", pt_ref.name.as_str()), ("tabId", tab_id.as_str())];
            xml_empty_tag(&mut self.writer, "x14:pivotTable", &attributes);
        }

        xml_end_tag(&mut self.writer, "x14:pivotTables");
    }

    /// Write the <x14:data> element.
    fn write_data(&mut self) {
        xml_start_tag(&mut self.writer, "x14:data", &[] as &[(&str, &str)]);

        self.write_tabular();

        xml_end_tag(&mut self.writer, "x14:data");
    }

    /// Write the <x14:tabular> element.
    fn write_tabular(&mut self) {
        let pivot_cache_id = self.pivot_cache_id.to_string();

        let attributes = [
            ("pivotCacheId", pivot_cache_id.as_str()),
            ("sortOrder", self.base.sort_order.to_sort_order_str()),
            ("crossFilter", self.base.cross_filter.to_cross_filter_str()),
            (
                "customListSort",
                if self.base.custom_list_sort { "1" } else { "0" },
            ),
        ];

        if self.items.is_empty() {
            xml_empty_tag(&mut self.writer, "x14:tabular", &attributes);
        } else {
            xml_start_tag(&mut self.writer, "x14:tabular", &attributes);

            self.write_items();

            xml_end_tag(&mut self.writer, "x14:tabular");
        }
    }

    /// Write the <x14:items> element.
    fn write_items(&mut self) {
        let count = self.items.len().to_string();
        let attributes = [("count", count.as_str())];

        xml_start_tag(&mut self.writer, "x14:items", &attributes);

        for item in self.items.clone() {
            self.write_item(&item);
        }

        xml_end_tag(&mut self.writer, "x14:items");
    }

    /// Write an <i> (item) element.
    fn write_item(&mut self, item: &SlicerItem) {
        let x = item.index.to_string();
        let mut attributes: Vec<(&str, &str)> = vec![("x", x.as_str())];

        if item.selected {
            attributes.push(("s", "1"));
        }

        if item.no_data {
            attributes.push(("nd", "1"));
        }

        xml_empty_tag(&mut self.writer, "i", &attributes);
    }
}

// -----------------------------------------------------------------------
// Slicer
// -----------------------------------------------------------------------

/// Represents an Excel slicer.
///
/// Slicers provide a visual way to filter table or pivot table data.
#[derive(Debug, Clone)]
pub struct Slicer {
    pub(crate) writer: Cursor<Vec<u8>>,

    /// Slicer name.
    pub(crate) name: String,

    /// Cache name (must match slicer cache).
    pub(crate) cache_name: String,

    /// Caption (display text).
    pub(crate) caption: String,

    /// Source type.
    pub(crate) source_type: SlicerSourceType,

    /// Row height in points.
    pub(crate) row_height: f64,

    /// Number of columns.
    pub(crate) column_count: u32,

    /// First visible item index.
    pub(crate) start_item: u32,

    /// Whether to show caption.
    pub(crate) show_caption: bool,

    /// Whether position is locked.
    pub(crate) locked_position: bool,

    /// Slicer style.
    pub(crate) style: SlicerStyle,

    /// Custom style name (if style is Other).
    pub(crate) style_name: String,
}

impl Slicer {
    /// Create a new slicer.
    ///
    /// # Arguments
    ///
    /// * `name` - The slicer name.
    /// * `cache_name` - The cache name (usually same as slicer name).
    /// * `caption` - The display caption.
    pub fn new(
        name: impl Into<String>,
        cache_name: impl Into<String>,
        caption: impl Into<String>,
    ) -> Slicer {
        let name = name.into();
        let cache_name = cache_name.into();
        let caption = caption.into();

        Slicer {
            writer: Cursor::new(Vec::with_capacity(2048)),
            name,
            cache_name,
            caption,
            source_type: SlicerSourceType::Table,
            row_height: 19.0,
            column_count: 1,
            start_item: 0,
            show_caption: true,
            locked_position: false,
            style: SlicerStyle::SlicerStyleLight4,
            style_name: String::new(),
        }
    }

    /// Set the source type.
    pub fn set_source_type(mut self, source_type: SlicerSourceType) -> Slicer {
        self.source_type = source_type;
        self
    }

    /// Set the row height in points.
    pub fn set_row_height(mut self, height: f64) -> Slicer {
        self.row_height = height;
        self
    }

    /// Set the number of columns.
    pub fn set_column_count(mut self, count: u32) -> Slicer {
        self.column_count = count;
        self
    }

    /// Set whether to show the caption.
    pub fn set_show_caption(mut self, show: bool) -> Slicer {
        self.show_caption = show;
        self
    }

    /// Set whether the position is locked.
    pub fn set_locked_position(mut self, locked: bool) -> Slicer {
        self.locked_position = locked;
        self
    }

    /// Set the slicer style.
    pub fn set_style(mut self, style: SlicerStyle) -> Slicer {
        self.style = style;
        self
    }

    /// Set a custom style name.
    pub fn set_style_name(mut self, name: impl Into<String>) -> Slicer {
        self.style_name = name.into();
        self.style = SlicerStyle::Other;
        self
    }
}

// -----------------------------------------------------------------------
// SlicerCollection
// -----------------------------------------------------------------------

/// A collection of slicers for a worksheet.
pub struct SlicerCollection {
    pub(crate) writer: Cursor<Vec<u8>>,
    pub(crate) slicers: Vec<Slicer>,
}

impl SlicerCollection {
    /// Create a new slicer collection.
    pub(crate) fn new() -> SlicerCollection {
        SlicerCollection {
            writer: Cursor::new(Vec::with_capacity(4096)),
            slicers: Vec::new(),
        }
    }

    /// Add a slicer to the collection.
    pub(crate) fn add_slicer(&mut self, slicer: Slicer) {
        self.slicers.push(slicer);
    }

    /// Check if the collection is empty.
    pub(crate) fn is_empty(&self) -> bool {
        self.slicers.is_empty()
    }

    /// Assemble the XML file for the slicer collection.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        self.write_slicers();

        xml_end_tag(&mut self.writer, "slicers");
    }

    /// Write the <slicers> element.
    fn write_slicers(&mut self) {
        let attributes = [
            (
                "xmlns",
                "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
            ),
            (
                "xmlns:mc",
                "http://schemas.openxmlformats.org/markup-compatibility/2006",
            ),
            ("mc:Ignorable", "x"),
            (
                "xmlns:x",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            ),
        ];

        xml_start_tag(&mut self.writer, "slicers", &attributes);

        for slicer in self.slicers.clone() {
            self.write_slicer(&slicer);
        }
    }

    /// Write a <slicer> element.
    fn write_slicer(&mut self, slicer: &Slicer) {
        let row_height = slicer.row_height.to_string();
        let column_count = slicer.column_count.to_string();
        let start_item = slicer.start_item.to_string();

        let style_name = if slicer.style == SlicerStyle::Other && !slicer.style_name.is_empty() {
            &slicer.style_name
        } else {
            slicer.style.to_style_str()
        };

        let mut attributes: Vec<(&str, &str)> = vec![
            ("name", &slicer.name),
            ("cache", &slicer.cache_name),
            ("caption", &slicer.caption),
        ];

        if slicer.show_caption {
            attributes.push(("showCaption", "1"));
        } else {
            attributes.push(("showCaption", "0"));
        }

        attributes.push(("rowHeight", row_height.as_str()));
        attributes.push(("columnCount", column_count.as_str()));
        attributes.push(("startItem", start_item.as_str()));

        if slicer.locked_position {
            attributes.push(("lockedPosition", "1"));
        } else {
            attributes.push(("lockedPosition", "0"));
        }

        attributes.push(("style", style_name));

        xml_empty_tag(&mut self.writer, "slicer", &attributes);
    }
}

impl Clone for SlicerCollection {
    fn clone(&self) -> Self {
        SlicerCollection {
            writer: Cursor::new(self.writer.get_ref().clone()),
            slicers: self.slicers.clone(),
        }
    }
}
