# OTR-Reports
script written to organize data from a large, weekly spreadsheet

The OTR is a spreadsheet sent from Nike/Jordan to my workplace, a shoe retailer, once a week. 
It details international container shipments. I have included two old OTRs (altered to protect my job) 
so that you can run the code yourself.

*Data Needed from OTR*

A) Deliveries that will arrive after the cancel date. The Cancel date is the date after which we don't 
want the product. Nike/Jordan shoes are incredibly time sensitive, and receiving them later than our competitors 
means we won't be able to sell them easily.

B) Deliveries that will arrive after the launch date, if applicable. Some shoes, especially those for Jordan, 
have a launch date. We are not allowed to sell the shoes before this date. Since these shoes often have a supply
that is purposefully less than the demand, the launch date is a big event. Having shoes arrive after the launch 
can cause us to lose sales and customers.

C) Launch date changes. Launch dates can and do change, with as little as a week's notice. The OTR, while not 
originally designed for this purpose, is the most reliable source of updates regarding launch date changes. It 
is also the most reliable list of launch dates that only includes shoes we intend to retail, as the launch calendars
available from Nike/Jordan includes all available shoes, of which we only retail a fraction.

*Launch Dates*

Jordan stopped supplying us wth launch dates on their OTR. To compensate, I downloaded the launch date for current and next season and added them to the process.

*To Run the Script*

1) Download the Nike OTR, Jordan OTR, and the macro-enabled workbook.

2) Open the macro-enabled workbook and fill out the fields. You must end the "Directory" field with a backslash
or the script won't be able to find your files. The "Today's Date" field will self-populate.

3) Ensure that sheet 2 is blank except for its header.

4) Press any button to run the associated report.

*Code Sources*

https://officetricks.com/different-methods-to-import-data-from-external-excel-file/
http://www.excel-easy.com
http://excelmacromastery.com/excel-vba-find/
https://www.techonthenet.com/excel/formulas/strcomp.php
https://msdn.microsoft.com/en-us/library/wts33hb3.aspx
