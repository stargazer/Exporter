from tablib import Dataset
from tablib.formats import xls
from tablib.compat import BytesIO, xlwt

def _clean(value):
    """
    Returns a "clean" and "serialized" representation of ``value``

    If value is a list or dict, clean recursively, and clean appropriately.
    """
    if isinstance(value, dict):
        return ', '.join([
            '%s: %s' % (_clean(k), _clean(v))
            for k, v in value.items()
        ])

    elif isinstance(value, list) \
        or isinstance(value, tuple) \
            or isinstance(value, set):
        return ', '.join([_clean(element) for element in value])

    try:
        return str(value)
    except UnicodeEncodeError:
        return unicode(value)

def export(data, fields, sheet_title=''):
    """
    @param data: Dictionary of {key: value} pairs, or list of dictionaries
    @param fields: The keys of the dictionaries that we want to export

    Serializes the data into an excel file, and returns the sheet
    """
    # if ``data`` is a dictionary, we make it into a list
    if isinstance(data, dict):
        data = [data, ]

    # We layout the values into a list of tuples, with each tuple element
    # corresponding to a header
    values = []
    for element in data:
        values.append(
            [_clean(element.get(key)) for key in fields]
        )

    headers = list(fields)

    # Detect formatting based on header values:
    #   **Bold**
    format_columns = {}
    for i, head in enumerate(headers):
        if head.startswith('**') and head.endswith('**'):
            format_columns[i] = 'B'
            headers[i] = head[2:-2]

    # Capitalize the headers
    headers = [header.capitalize() for header in headers]

    # Create the tablib dataset
    dataset = Dataset(*values, headers=headers, title=sheet_title)

    # Create worksheet from dataset
    workbook = xlwt.Workbook(encoding='utf8')
    worksheet = workbook.add_sheet(dataset.title, cell_overwrite_ok=True)
    xls.dset_sheet(dataset, worksheet)  # actual conversion

    # Overwrite columns with formatted cells
    bold = xlwt.easyxf("font: bold on")  #TODO Change this into Font()
    for row in range(dataset.height):
        for col, style in format_columns.iteritems():
            if 'B' in style:
                # `row+1`, since we want to pass over the header row
                worksheet.write(row+1, col, dataset[row][col], bold)

    # Export to excel
    stream = BytesIO()
    workbook.save(stream)
    stream.seek(0)

    return stream
