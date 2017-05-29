dodotable-xlsx
==============

.. image:: https://badge.fury.io/py/dodotable-xlsx.svg
   :target: https://pypi.python.org/pypi/dodotable-xlsx
   :alt: Latest PyPI version

Excel (.xlsx) exporter for dodotable_.

.. _dodotable: https://github.com/spoqa/dodotable


Changelog
---------

Version 0.2.0
~~~~~~~~~~~~~

To be released.

- ``datetime.date`` values became properly formatted to ``yyyy-mm-dd``.

  - ``dodotable_xlsx.write_table_to_workbook()`` function became to have
    an optional parameter ``date_format``.

  - ``dodotable_xlsx.write_table_to_worksheet()`` function became to have
    a required parameter ``date_format``.


Version 0.1.0
~~~~~~~~~~~~~

Released on May 25, 2017.
