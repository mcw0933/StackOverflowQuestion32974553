# StackOverflowQuestion32974553
Sample files for [StackOverflow question 32974553](http://stackoverflow.com/questions/32974553/how-can-i-post-process-the-data-from-an-excel-web-query-when-the-query-is-comple/32975304?noredirect=1#comment53781999_32975304)

Contains an .xlsm Excel workbook with two tabs.  

Two buttons in the first tab will both trigger a web query that populates the second tab.

In order to catch the completion of that load, the one button hooks up an event handler to the QueryTable.  The other button does not.

Source code for the workbook is in two VBA files in the src folder.