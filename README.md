# xls2sql

A tool for automated import of spreadsheet-data

```
python3 xls2db.py -i1 -c2 -f3 basic.xlsx | mysql --user=test --password=password basic
mysql --user=test --password=password basic -e "select * from Sheet1;"
```

