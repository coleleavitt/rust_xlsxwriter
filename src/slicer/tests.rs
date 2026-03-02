// Slicer unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod slicer_tests {
    use crate::slicer::{
        PivotTableSlicerCache, PivotTableSlicerRef, Slicer, SlicerCrossFilter, SlicerItem,
        SlicerSortOrder, SlicerSourceType, SlicerStyle, TableSlicerCache,
    };

    #[test]
    fn test_slicer_source_type() {
        let slicer =
            Slicer::new("Test", "Test", "Test Caption").set_source_type(SlicerSourceType::Table);
        assert_eq!(slicer.source_type, SlicerSourceType::Table);

        let slicer = Slicer::new("Test", "Test", "Test Caption")
            .set_source_type(SlicerSourceType::PivotTable);
        assert_eq!(slicer.source_type, SlicerSourceType::PivotTable);
    }

    #[test]
    fn test_slicer_cross_filter() {
        assert_eq!(SlicerCrossFilter::None.to_cross_filter_str(), "none");
        assert_eq!(
            SlicerCrossFilter::ShowItemsWithDataAtTop.to_cross_filter_str(),
            "showItemsWithDataAtTop"
        );
        assert_eq!(
            SlicerCrossFilter::ShowItemsWithNoData.to_cross_filter_str(),
            "showItemsWithNoData"
        );
    }

    #[test]
    fn test_slicer_sort_order() {
        assert_eq!(SlicerSortOrder::Ascending.to_sort_order_str(), "ascending");
        assert_eq!(
            SlicerSortOrder::Descending.to_sort_order_str(),
            "descending"
        );
    }

    #[test]
    fn test_slicer_style() {
        assert_eq!(
            SlicerStyle::SlicerStyleLight1.to_style_str(),
            "SlicerStyleLight1"
        );
        assert_eq!(
            SlicerStyle::SlicerStyleLight4.to_style_str(),
            "SlicerStyleLight4"
        );
        assert_eq!(
            SlicerStyle::SlicerStyleDark1.to_style_str(),
            "SlicerStyleDark1"
        );
        assert_eq!(
            SlicerStyle::SlicerStyleDark6.to_style_str(),
            "SlicerStyleDark6"
        );
    }

    #[test]
    fn test_slicer_item() {
        let item = SlicerItem::new(0);
        assert_eq!(item.index, 0);
        assert!(item.selected);
        assert!(!item.no_data);

        let item = SlicerItem::new(5).set_selected(false).set_no_data(true);
        assert_eq!(item.index, 5);
        assert!(!item.selected);
        assert!(item.no_data);
    }

    #[test]
    fn test_slicer_creation() {
        let slicer = Slicer::new("Slicer_Region", "Slicer_Region", "Region");

        assert_eq!(slicer.name, "Slicer_Region");
        assert_eq!(slicer.cache_name, "Slicer_Region");
        assert_eq!(slicer.caption, "Region");
        assert!(slicer.show_caption);
        assert!(!slicer.locked_position);
        assert_eq!(slicer.column_count, 1);
    }

    #[test]
    fn test_slicer_options() {
        let slicer = Slicer::new("Test", "Test", "Caption")
            .set_row_height(25.0)
            .set_column_count(3)
            .set_show_caption(false)
            .set_locked_position(true)
            .set_style(SlicerStyle::SlicerStyleDark2);

        assert_eq!(slicer.row_height, 25.0);
        assert_eq!(slicer.column_count, 3);
        assert!(!slicer.show_caption);
        assert!(slicer.locked_position);
        assert_eq!(slicer.style, SlicerStyle::SlicerStyleDark2);
    }

    #[test]
    fn test_slicer_custom_style() {
        let slicer = Slicer::new("Test", "Test", "Caption").set_style_name("MyCustomStyle");

        assert_eq!(slicer.style, SlicerStyle::Other);
        assert_eq!(slicer.style_name, "MyCustomStyle");
    }

    #[test]
    fn test_table_slicer_cache() {
        let cache = TableSlicerCache::new("Slicer_Region", "Region", 1, 1)
            .set_sort_order(SlicerSortOrder::Descending)
            .set_cross_filter(SlicerCrossFilter::ShowItemsWithDataAtTop);

        assert_eq!(cache.base.name, "Slicer_Region");
        assert_eq!(cache.base.source_name, "Region");
        assert_eq!(cache.table_id, 1);
        assert_eq!(cache.column_id, 1);
        assert_eq!(cache.base.sort_order, SlicerSortOrder::Descending);
        assert_eq!(
            cache.base.cross_filter,
            SlicerCrossFilter::ShowItemsWithDataAtTop
        );
    }

    #[test]
    fn test_pivot_table_slicer_cache() {
        let cache = PivotTableSlicerCache::new("Slicer_Region", "Region", 1)
            .add_pivot_table(PivotTableSlicerRef::new("PivotTable1", 1))
            .add_item(SlicerItem::new(0))
            .add_item(SlicerItem::new(1).set_selected(false));

        assert_eq!(cache.base.name, "Slicer_Region");
        assert_eq!(cache.pivot_cache_id, 1);
        assert_eq!(cache.pivot_tables.len(), 1);
        assert_eq!(cache.items.len(), 2);
    }

    #[test]
    fn test_pivot_table_slicer_ref() {
        let pt_ref = PivotTableSlicerRef::new("MyPivot", 2);
        assert_eq!(pt_ref.name, "MyPivot");
        assert_eq!(pt_ref.tab_id, 2);
    }
}
