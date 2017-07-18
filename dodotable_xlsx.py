import datetime
import logging

from dodotable.schema import LinkedCell, Table
from markupsafe import Markup
from xlsxwriter import Workbook
from xlsxwriter.format import Format
from xlsxwriter.worksheet import Worksheet

__version__ = '0.3.0'
__all__ = 'write_table_to_workbook', 'write_table_to_worksheet'


def write_table_to_workbook(table, workbook,
                            header_format=None,
                            date_format=None,
                            chunk_size=None):
    if not isinstance(table, Table):
        raise TypeError(
            'table must be an instance of {0.__module__}.{0.__name__} or its '
            'subclass, not {1!r}'.format(Table, table)
        )
    elif not isinstance(workbook, Workbook):
        raise TypeError('workbook must be an instance of {0.__module__}.'
                        '{0.__name__}, not {1!r}'.format(Workbook, workbook))
    worksheet = workbook.add_worksheet(table.label)
    if header_format is None:
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': 'black',
            'fg_color': 'white',
        })
    if date_format is None:
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    write_table_to_worksheet(table, worksheet,
                             header_format=header_format,
                             date_format=date_format, chunk_size=chunk_size)


def write_table_to_worksheet(table, worksheet,
                             header_format,
                             date_format,
                             chunk_size):
    logger = logging.getLogger(__name__ + '.write_table_to_worksheet')
    if not isinstance(table, Table):
        raise TypeError(
            'table must be an instance of {0.__module__}.{0.__name__} or its '
            'subclass, not {1!r}'.format(Table, worksheet)
        )
    elif not isinstance(worksheet, Worksheet):
        raise TypeError('worksheet must be an instance of {0.__module__}.'
                        '{0.__name__}, not {1!r}'.format(Worksheet, worksheet))
    elif not isinstance(header_format, Format):
        raise TypeError(
            'header_format must be an instance of {0.__module__}.{0.__name__},'
            ' not {1!r}'.format(Format, header_format)
        )
    for col, column in enumerate(table.columns):
        worksheet.write(0, col, column.label, header_format)
        logger.debug('Column #%d: %s', col, column.label)
    logger.debug('%s', table.query)

    count = table.count
    offset = 0
    if chunk_size is None:
        # rows may become more
        chunk_size = count + 100

    rn = 0
    while offset <= count:
        table.select(offset=offset, limit=chunk_size)
        logger.debug('Fetch from %d to %d', offset, chunk_size)
        offset += chunk_size

        for row in table.rows:
            rn += 1
            logger.debug('Row number: %d', rn)
            for col, cell in enumerate(row):
                val = cell.data
                for mode in 'without_render', 'with_render':
                    # If a cell.data is a plain Python value (e.g. int, str)
                    # these can be adapted by XlsxWriter's renderer.
                    # However, if it's a complext value these need to be
                    # rendered by dodotable prior to be passed to XlsxWriter.
                    try:
                        if isinstance(cell, LinkedCell):
                            worksheet.write_url(rn, col, cell.url,
                                                string=str(val))
                        elif (isinstance(val, datetime.date) and
                              not isinstance(val, datetime.datetime)):
                            worksheet.write_datetime(rn, col, val, date_format)
                        else:
                            worksheet.write(rn, col, val)
                    except TypeError:
                        if mode == 'without_render':
                            val = cell.repr(val)
                            if isinstance(val, Markup):
                                val = val.striptags()
                            continue
                        raise
                    else:
                        break
