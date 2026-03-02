// Pivot table unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod pivot_table_tests {
    use crate::pivot_table::{
        DataFieldFunction, DateGroupBy, PivotDataField, PivotField, PivotFieldAxis, PivotFieldItem,
        PivotTable, ShowDataAs,
    };

    #[test]
    fn test_pivot_field_creation() {
        let field = PivotField::new("Region");
        assert_eq!(field.name, "Region");
        assert_eq!(field.axis, PivotFieldAxis::None);
        assert!(!field.data_field);
    }

    #[test]
    fn test_pivot_field_axis() {
        let field = PivotField::new("Category").set_axis(PivotFieldAxis::Row);
        assert_eq!(field.axis, PivotFieldAxis::Row);
        assert_eq!(field.axis.to_axis_str(), Some("axisRow"));
    }

    #[test]
    fn test_pivot_field_axis_strings() {
        assert_eq!(PivotFieldAxis::None.to_axis_str(), None);
        assert_eq!(PivotFieldAxis::Row.to_axis_str(), Some("axisRow"));
        assert_eq!(PivotFieldAxis::Column.to_axis_str(), Some("axisCol"));
        assert_eq!(PivotFieldAxis::Page.to_axis_str(), Some("axisPage"));
        assert_eq!(PivotFieldAxis::Values.to_axis_str(), Some("axisValues"));
    }

    #[test]
    fn test_data_field_function() {
        let data_field =
            PivotDataField::new("Sum of Sales", 1).set_function(DataFieldFunction::Sum);

        assert_eq!(data_field.name, "Sum of Sales");
        assert_eq!(data_field.field_index, 1);
        assert_eq!(data_field.function, DataFieldFunction::Sum);
    }

    #[test]
    fn test_data_field_functions() {
        assert_eq!(DataFieldFunction::Sum.to_subtotal_str(), "sum");
        assert_eq!(DataFieldFunction::Count.to_subtotal_str(), "count");
        assert_eq!(DataFieldFunction::Average.to_subtotal_str(), "average");
        assert_eq!(DataFieldFunction::Max.to_subtotal_str(), "max");
        assert_eq!(DataFieldFunction::Min.to_subtotal_str(), "min");
        assert_eq!(DataFieldFunction::Product.to_subtotal_str(), "product");
        assert_eq!(DataFieldFunction::CountNums.to_subtotal_str(), "countNums");
        assert_eq!(DataFieldFunction::StdDev.to_subtotal_str(), "stdDev");
        assert_eq!(DataFieldFunction::StdDevP.to_subtotal_str(), "stdDevp");
        assert_eq!(DataFieldFunction::Var.to_subtotal_str(), "var");
        assert_eq!(DataFieldFunction::VarP.to_subtotal_str(), "varp");
    }

    #[test]
    fn test_show_data_as() {
        let data_field =
            PivotDataField::new("% of Total", 1).set_show_data_as(ShowDataAs::PercentOfTotal);

        assert_eq!(data_field.show_data_as, ShowDataAs::PercentOfTotal);
        assert_eq!(
            data_field.show_data_as.to_show_data_as_str(),
            Some("percentOfTotal")
        );
    }

    #[test]
    fn test_show_data_as_strings() {
        assert_eq!(ShowDataAs::Normal.to_show_data_as_str(), None);
        assert_eq!(
            ShowDataAs::Difference.to_show_data_as_str(),
            Some("difference")
        );
        assert_eq!(
            ShowDataAs::PercentDifference.to_show_data_as_str(),
            Some("percentDiff")
        );
        assert_eq!(
            ShowDataAs::PercentOfRow.to_show_data_as_str(),
            Some("percentOfRow")
        );
        assert_eq!(
            ShowDataAs::PercentOfColumn.to_show_data_as_str(),
            Some("percentOfCol")
        );
        assert_eq!(
            ShowDataAs::PercentOfTotal.to_show_data_as_str(),
            Some("percentOfTotal")
        );
        assert_eq!(ShowDataAs::Index.to_show_data_as_str(), Some("index"));
        assert_eq!(
            ShowDataAs::RunningTotal.to_show_data_as_str(),
            Some("runTotal")
        );
        assert_eq!(
            ShowDataAs::RankAscending.to_show_data_as_str(),
            Some("rankAscending")
        );
        assert_eq!(
            ShowDataAs::RankDescending.to_show_data_as_str(),
            Some("rankDescending")
        );
    }

    #[test]
    fn test_pivot_field_item() {
        let item = PivotFieldItem::new(0).set_hidden(true).set_type("default");

        assert_eq!(item.index, 0);
        assert!(item.hidden);
        assert_eq!(item.item_type, Some("default".to_string()));
    }

    #[test]
    fn test_pivot_table_creation() {
        let pivot = PivotTable::new()
            .set_name("PivotTable1")
            .set_data_range("Sheet1!A1:D100");

        assert_eq!(pivot.name, "PivotTable1");
        assert_eq!(pivot.data_range, "Sheet1!A1:D100");
        assert!(pivot.row_grand_totals);
        assert!(pivot.col_grand_totals);
    }

    #[test]
    fn test_pivot_table_fields() {
        let pivot = PivotTable::new()
            .set_name("Test")
            .add_row_field(PivotField::new("Region"))
            .add_column_field(PivotField::new("Quarter"))
            .add_page_field(PivotField::new("Year"))
            .add_data_field(PivotDataField::new("Sum of Sales", 3));

        assert_eq!(pivot.row_fields.len(), 1);
        assert_eq!(pivot.column_fields.len(), 1);
        assert_eq!(pivot.page_fields.len(), 1);
        assert_eq!(pivot.data_fields.len(), 1);

        assert_eq!(pivot.row_fields[0].axis, PivotFieldAxis::Row);
        assert_eq!(pivot.column_fields[0].axis, PivotFieldAxis::Column);
        assert_eq!(pivot.page_fields[0].axis, PivotFieldAxis::Page);
    }

    #[test]
    fn test_pivot_table_grand_totals() {
        let pivot = PivotTable::new()
            .set_row_grand_totals(false)
            .set_col_grand_totals(false);

        assert!(!pivot.row_grand_totals);
        assert!(!pivot.col_grand_totals);
    }

    #[test]
    fn test_pivot_table_layout() {
        let pivot = PivotTable::new().set_compact(false).set_outline(true);

        assert!(!pivot.compact);
        assert!(pivot.outline);
    }

    #[test]
    fn test_pivot_table_data_caption() {
        let pivot = PivotTable::new().set_data_caption("Values");

        assert_eq!(pivot.data_caption, "Values");
    }
}
