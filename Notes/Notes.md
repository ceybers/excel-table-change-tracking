# Notes

0. Click "Start Session"
1. Copy tankdb -> arr
2. Create bufferBefore worksheet
3. Copy arr -> bufferBefore
4. Wait for user to click "End Session" (or "Cancel Session")
5. Copy bufferBefore -> arrLHS
6. Copy tankdb -> arrRHS (no need for bufferAfter worksheet)
7. Create arrBitmask
8. Compare with (i, j) for-loop
9. Create collection for class of ~~tuples~~ changes
10. For each 1 in arrBitmask, add a new tuple: Enitity (tank code), Attribute (column name), ValueBefore, ValueAfter
11. Write this collection to a Changelog worksheet, or to a multi-line textbox in a dialog, or straight to an open MS Access instance
12. Check that all the keys are valid (no new added rows)
13. Check that all the columns are mapped: Excel.ColumnName -> tblDetail.FieldName
14. Group them together by tankCode×tblDetail (instead of tankCode×tblDetail.FieldName)
15. Check if existing record exists. If yes, update tblTrack.ValidFrom from 9999 to now()
16. Create new temporary tblUpdRef* (ask if 1x for all, or 1x tank)
17. Create 1x tblTrack and 1x tblDetail row for each tankCode×tblDetail pair
18. Finished!

working (main tankdb table)
repo (original pre-edit)
staging (post-edit before commit)

pull: replace working with repo (undo all changes)
do work on tankdb
add: updates staging with rows/cols/cells
commit: calculates diff between repo and staging, then commits them to staging (and appends to change log)
at this point after commit, working == staging == repo (if all changes in working were added)

overcomplicated.

working
before arr (store as static, might need as WS in case VBA instance bugs out)
after  arr (not required as WS, can be done in memory)
changes (flat list)