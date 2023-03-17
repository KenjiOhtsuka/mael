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

This tool is installed with pip::

    $ pip install mael
    
*****
Usage
*****

1. Initialize the directory, then initial files are created based on the templates.::

        $ mael init some_dir
        
2. Write your data in markdown. You can put multiple markdown files in the directory.::

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
        
3. Build Excel, then you can get an Excel file in the directory.::

        $ mael build some_dir
        
There, the Excel file is as:

    Summary
    
    Please write summary of the table data.

    +-----------+-----------+-----------+
    | Column 1  | Column 2  | Column 3  |
    +-----------+-----------+-----------+
    | Value 1-1 | Value 2-1 |           |
    +-----------+-----------+-----------+
    | Value 2-1 |           | Value 3-2 |
    +-----------+-----------+-----------+
    
If you put multiple markdown files, the Excel file contains multiple sheets.
