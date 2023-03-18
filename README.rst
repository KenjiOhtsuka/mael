====
mael
====

Tool to convert **ma**\ rkdown list to Exc\ **el**, mael.

**********
Motivation
**********

Do you use Excel for summarizing some lists as tables?
Isn't it difficult to manage Excel in git?
The common repositories show differences in text files between versions but not in Excel files.
How can we manage spreadsheet tables with versions?
CSV is one of the choices but is it really easy?
What if we can manage table data as markdown file?

This tool can convert markdown files to tables in an Excel file.

************
Installation
************

This tool is installed with pip:

.. code-block:: bash

    $ pip install mael

*****
Usage
*****

#. Initialize the directory, then initial files are created based on the templates.

   .. code-block:: bash

     $ mael init some_dir

#. Write your data in markdown. You can put multiple markdown files in the directory.

   .. code-block:: markdown

     # List title

     ## Summary

     Please write summary of the table data.

     ## List

     ### Column 1

     Value 1-1

     ### Column 2

     Value 1-2

     ---

     ### Column 1

     Value 2-1

     ---

     ### Column 2

     Value 3-2

   Separate each item with :code:`---`.

#. Build Excel, then you can get an Excel file in the directory.

   .. code-block:: bash

     $ mael build some_dir

   There, the Excel file contains the sheet as:

     **Summary**

     Please write summary of the table data.

     +-----------+-----------+-----------+
     | Column 1  | Column 2  | Column 3  |
     +-----------+-----------+-----------+
     | Value 1-1 | Value 2-1 |           |
     +-----------+-----------+-----------+
     | Value 2-1 |           | Value 3-2 |
     +-----------+-----------+-----------+

   If you put multiple markdown files, the Excel file contains multiple sheets.

************
Advanced use
************

You can use variables.
Also, you can define environmental variables for each environment.

#. Define variables in :code:`some_dir/config/variables.ini`:

   .. code-block:: ini

     VARIABLE_1=ABCDEFG
     VARIABLE_2=HIJKLMN

#. Use the variables in markdown files.
   Surround the variable name with :code:`{{` and :code:`}}`:

   .. code-block:: markdown

     # List title

     ## Summary

     Variable 1 is {{ VARIABLE_1 }}.
     Variable 2 is {{ VARIABLE_2 }}.

     ......

   Of course, you can use the variables not only in the summary but also in the list.

#. Build Excel, then you can get an Excel file in the directory.

   .. code-block:: bash

     $ mael build some_dir

   There, the Excel file contains the sheet as:

     **Summary**

     | variable 1 is ABCDEFG.
     | variable 2 is HIJKLMN.

     | \.\.\.\.\.\.

To use environmental variables, define the variables in :code:`some_dir/config/variables.${env_name}.ini`, such as :code:`some_dir/config/variables.dev.ini`. Environmental variable file overwrite the varabiles defined in the normal variable file, :code:`variable.ini`. To build the environmental file, execute :code:`mael build some_dir -e dev`, and you will get the Excel file, :code:`some_dir_dev.xlsx`.

************
PyPI package
************

https://pypi.org/project/mael/
