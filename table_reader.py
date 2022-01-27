# module to extract data from a spreadsheet with known label addresses
key_map = {
    'pinYear': 'A',
    'pin': 'B',
    'pinCost': 'C',
}


def fetch_expected_specs_in_(c_sheet, active_row: int) -> dict:
    """
    :param c_sheet: a sheet from a workbook object
    :param active_row: an integer corresponding to a spreadsheet row
    :return: a dictionary containing one row of data
    """
    pie_specs = {}
    for expected_key in key_map:
        expected_value = c_sheet[key_map[expected_key] +
                                 str(active_row)].value
        pie_specs.update({
            expected_key: expected_value
        })
    return pie_specs
