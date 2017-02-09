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

*To Run the Script*

1) Download the Nike OTR, Jordan OTR, and the macro-enabled workbook.

2) Open the macro-enabled workbook and fill out the top three fields. You must end the "Directory" field with a backslash
or the script won't be able to find your files. The "Today's Date" field will self-populate.

3) Press any button to run the associated report.

*Output*

Cancel Dates: This report prints out on Sheet2. The output is a list of Sales Order numbers. Orders highlighted in Blue 
are expected to arrive past the Cancel Date. Orders highlighted in pink are expected to arrive less than a week before 
the launch date. If an order qualifies for both, it will appear twice, once in blue and once in pink.

Launch: This report prints out to Sheet3. The output is simply our internal Purchase Order number, the Material number, 
and Launch date of any order which has a launch date listed.

*Code Sources*

The project started by copying this person's code for how to:

https://officetricks.com/different-methods-to-import-data-from-external-excel-file/

And the rest was put together using this tutorial:

http://www.excel-easy.com
