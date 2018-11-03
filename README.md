# xls2sql

A tool for automated import of spreadsheet-data

```
python3 xls2db.py --drop -i1 -c2 -f3 basic.xlsx | mysql --user=test --password=password basic
mysql --user=test --password=password basic -e "select * from Sheet1;"

python3 xls2sql.py --drop -l'2-3,4,2' -i1 -c2 -f3 basic.xlsx
python3 xls2sql.py --drop -i1 -c2 -f3 missing.xlsx 
```

