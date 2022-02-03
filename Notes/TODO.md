# TO-DO

- [*] Basic implementation
- [*] Check for orphaned keys
- [*] Check for orphaned columns
- [*] Schema to map Excel.ListObject to Access.Tables
- [*] Grouping Excel fields by Access tables
- [*] Insert new record into SCD2 model
- [*] Update existing and insert amended record into SCD2 model
- [*] Commit strategy: per session, per key
- [ ] Change header formatting from .Interior.Color to ConditionalFormatting where =TRUE
- [*] "Number Stored as Text" is VarType(8) on Working but VarType(5) on BeforeWS
- [*] Accomodate Lookup Columns on Access DB
- [ ] Need to handle 1:m relations, e.g. MaintHistoryLatest -> MaintFistoryFEI -> picks up ValidFrom and Reference from tblCommits
      Might need to downgrade to separate simple tables with implicit fields for date and refs
- [ ] Consider loading table and field names from meta* tables from schema
- [ ] Consider replacing const strings with a Collection
- [ ] User Interface
- [ ] Get list of orphaned keys
- [ ] Get list of orphaned columns
- [ ] Ignore list for columns
- [ ] Revert to Before state (Undo all changes)
- [ ] Commit all to database (keyframe vs delta frame)
