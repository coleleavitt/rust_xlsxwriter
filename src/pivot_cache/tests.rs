// Pivot cache unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod pivot_cache_tests {
    use crate::pivot_cache::{
        PivotCacheDefinition, PivotCacheField, PivotCacheRecord, PivotCacheRecordValue,
        PivotCacheRecords, PivotCacheSourceType, PivotCacheValue,
    };

    #[test]
    fn test_cache_source_type() {
        assert_eq!(PivotCacheSourceType::Worksheet.to_type_str(), "worksheet");
        assert_eq!(PivotCacheSourceType::External.to_type_str(), "external");
        assert_eq!(
            PivotCacheSourceType::Consolidation.to_type_str(),
            "consolidation"
        );
    }

    #[test]
    fn test_cache_field_creation() {
        let field = PivotCacheField::new("Region");
        assert_eq!(field.name, "Region");
        assert!(field.shared_items.is_empty());
        assert!(!field.contains_string);
        assert!(!field.contains_number);
    }

    #[test]
    fn test_cache_field_with_strings() {
        let field = PivotCacheField::new("Region")
            .add_string("North")
            .add_string("South")
            .add_string("East")
            .add_string("West");

        assert_eq!(field.shared_items.len(), 4);
        assert!(field.contains_string);
        assert!(!field.contains_number);
    }

    #[test]
    fn test_cache_field_with_numbers() {
        let field = PivotCacheField::new("Sales")
            .add_number(100.0)
            .add_number(250.5)
            .add_number(50.0);

        assert_eq!(field.shared_items.len(), 3);
        assert!(field.contains_number);
        assert_eq!(field.min_value, Some(50.0));
        assert_eq!(field.max_value, Some(250.5));
    }

    #[test]
    fn test_cache_field_with_integers() {
        let field = PivotCacheField::new("Quantity")
            .add_number(10.0)
            .add_number(20.0)
            .add_number(30.0);

        assert!(field.contains_integer);
    }

    #[test]
    fn test_cache_field_with_missing() {
        let field = PivotCacheField::new("Optional")
            .add_string("Value")
            .add_missing();

        assert!(field.contains_blank);
        assert_eq!(field.shared_items.len(), 2);
    }

    #[test]
    fn test_cache_record_creation() {
        let record = PivotCacheRecord::new()
            .add_shared_index(0)
            .add_number(100.0)
            .add_string("Q1");

        assert_eq!(record.values.len(), 3);
    }

    #[test]
    fn test_cache_record_values() {
        let record = PivotCacheRecord::new()
            .add_shared_index(5)
            .add_number(123.45)
            .add_boolean(true)
            .add_datetime("2024-01-15")
            .add_missing();

        assert_eq!(record.values.len(), 5);

        match &record.values[0] {
            PivotCacheRecordValue::SharedItemIndex(i) => assert_eq!(*i, 5),
            _ => panic!("Expected SharedItemIndex"),
        }

        match &record.values[1] {
            PivotCacheRecordValue::Number(n) => assert_eq!(*n, 123.45),
            _ => panic!("Expected Number"),
        }
    }

    #[test]
    fn test_cache_definition_creation() {
        let cache = PivotCacheDefinition::new()
            .set_source_range("Sheet1", "A1:D100")
            .set_record_count(99);

        assert_eq!(cache.source_sheet, "Sheet1");
        assert_eq!(cache.source_ref, "A1:D100");
        assert_eq!(cache.record_count, 99);
        assert_eq!(cache.source_type, PivotCacheSourceType::Worksheet);
    }

    #[test]
    fn test_cache_definition_with_table() {
        let cache = PivotCacheDefinition::new()
            .set_source_range("Sheet1", "A1:D100")
            .set_source_table("SalesTable");

        assert_eq!(cache.source_table, Some("SalesTable".to_string()));
    }

    #[test]
    fn test_cache_definition_with_fields() {
        let cache = PivotCacheDefinition::new()
            .add_field(
                PivotCacheField::new("Region")
                    .add_string("North")
                    .add_string("South"),
            )
            .add_field(
                PivotCacheField::new("Sales")
                    .add_number(100.0)
                    .add_number(200.0),
            );

        assert_eq!(cache.fields.len(), 2);
        assert_eq!(cache.fields[0].name, "Region");
        assert_eq!(cache.fields[1].name, "Sales");
    }

    #[test]
    fn test_cache_definition_options() {
        let cache = PivotCacheDefinition::new()
            .set_save_data(false)
            .set_refresh_on_load(true);

        assert!(!cache.save_data);
        assert!(cache.refresh_on_load);
    }

    #[test]
    fn test_cache_records_creation() {
        let records = PivotCacheRecords::new()
            .add_record(
                PivotCacheRecord::new()
                    .add_shared_index(0)
                    .add_number(100.0),
            )
            .add_record(
                PivotCacheRecord::new()
                    .add_shared_index(1)
                    .add_number(200.0),
            );

        assert_eq!(records.count(), 2);
    }

    #[test]
    fn test_cache_records_empty() {
        let records = PivotCacheRecords::new();
        assert_eq!(records.count(), 0);
    }
}
