FieldChange: Key, Column, Before, After

DatabaseChange: KeyFK, TrackFK, TableName, FieldName, After

KeyTranslation: ID, Key

SchemaItem: ColumnName, TableName, FieldName, VarType

FieldChange + SchemaItem + KeyTranslation => DatabaseChange

GroupedChanges: KeyFK, TableName, Collection<DatabaseChange>

GroupedChanges + TrackStrategy => Collection<Tracks>, and GroupedChanges updated with .TrackFK