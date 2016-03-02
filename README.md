# Exporter

Exports python data structures to other file formats. For the time being only
works for exporting to excel files.

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

## Formatting

#### Spaces in column headers
All underscores in the column headers are replaced by spaces. Example:
``First_name`` becomes ``First name``.


#### Capitalized column headers
All column headers are capitalized. This means that every letter except the
first is converted to lower case. The first character is converted to upper
case. Example: ``name`` becomes ``Name``.


#### Bold columns
If you want to use formatting in your columns, you can make this known to the
exporter know this by changing the column header names. So for a bold column
``name``, you name the column header ``**name**``. Then ``**name**`` is renamed
to ``name`` and every value in that column (except the column header itself) is
made bold.

Currently supported formats are: ``**bold**``
