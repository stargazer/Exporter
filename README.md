# Exporter

Exports python data structures to other file formats. For the time being only works
for exporting to excel files.

## Usage

```python
from exporter import excel

# Data to export
data = [
    {'id': 1, 'name': 'Randy', 'surname': 'Marsh'},
    {'id': 2, 'name': 'Eric', 'surname': 'Cartman', 'age': 8},
    {'id': 3, 'name': 'Herbert', 'surname': 'Garrison'},
]
# Selection of fields to export
fields = ['id', 'name', 'surname']

excel.export(data, fields, sheet_name='Export')
```
