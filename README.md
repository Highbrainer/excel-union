# ExcelUnion

A simplisitic utility to merge two excel files, based on a common "key" column.

Does something very similar to 
```sql
select * from file1 full outer join file2 on file1.keyCol1 = file2.keyCol2
```
