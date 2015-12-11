import tablib
import StringIO

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

    # Capitalize the headers        
    headers = [field.capitalize() for field in fields]        

    # Create the tablib dataset
    dataset = tablib.Dataset(*values, headers=headers, title=sheet_title)
    # Export to excel
    stream = StringIO.StringIO()
    stream.write(dataset.xls)
    stream.seek(0)

    return stream    
