dodotable-xlsx
==============

.. image:: https://badge.fury.io/py/dodotable-xlsx.svg
   :target: https://pypi.python.org/pypi/dodotable-xlsx
   :alt: Latest PyPI version

Excel (.xlsx) exporter for dodotable_.

.. _dodotable: https://github.com/spoqa/dodotable


Changelog
---------

Version 0.3.0
~~~~~~~~~~~~~

Released on July 18, 2017.

- Added ``chunk_size`` parameter on ``write_table_to_workbook()`` and
  ``write_table_to_worksheet``.


Version 0.2.0
~~~~~~~~~~~~~

Released on May 29, 2017.

- ``datetime.date`` values became properly formatted to ``yyyy-mm-dd``.

  - ``dodotable_xlsx.write_table_to_workbook()`` function became to have
    an optional parameter ``date_format``.

  - ``dodotable_xlsx.write_table_to_worksheet()`` function became to have
    a required parameter ``date_format``.

- Fixed a bug that ``dodotable_xlsx.write_table_to_workbook()`` function had
  raised ``NameError`` when the given ``table`` is an instance of an invalid
  type.


Version 0.1.0
~~~~~~~~~~~~~

Released on May 25, 2017.
