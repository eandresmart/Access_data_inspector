# Access_data_inspector

Have you got new data you are unfamiliar with and want to quickly build a picture of what you have? This tool might help you. The data does not even need to be from an OFM project - you can use it on any MS Access database, not just oilfield data.

**Steps to Use the Tool:**

1️⃣ **Run the ‘CheckAccessFile’ Procedure**: This will allow you to select an MS Access database file and generate an Excel table listing the tables in that file and their fields. You can then filter the tables and fields that may be of interest to you using the filter features in Excel.

2️⃣ **Run the ‘PopulateData’ Procedure**: This will give you some statistics about the data, including record count, how many of those records have non-null values, and average, maximum, minimum, and percentiles for the tables and fields filtered.

**How is this useful?**

You can quickly spot:

✔ Data tables and fields that are defined in the database but have no data.

✔ At a glance, you can see the sort of values you can expect, e.g., in terms of historical data period, average rates and pressures, etc.

✔ The extreme values (maximum, minimum, and even some percentiles) will likely point out any abnormalities in the data, to guide further data validation and cleaning. For example, daily data that has 48 hours of uptime, or negative production volumes - it happens in the best companies.


You will notice that I have percentiles for text fields - I would rather see some data than not, and although mathematically percentiles are not defined for text, computers have no issue in sorting strings in order. Based on that, I retrieve the percentile values: all field values are sorted, and the percentiles looked up treating them like an empirical distribution.

**IMPORTANT**

The procedures use DAO (Data Access Objects) objects and methods, so you need a reference to a library that will allow you to use them. For those familiar with Python this is similar to importing a library in a script - but it is less conveniently done in VBA.

The reference is already defined in the spreadsheet I'm sharing, but the inconvenience with VBA is that the reference depends on the path to your MS Office installation, and this may vary between PCs. So, if it is not working, most likely the existing reference does not work for your PC and you need to create it.

Creating the reference is done from the VBA IDE, and as it is I'm using the 'Microsoft Office 16.0 Access Database Engine Object Library', provided by the ACEDAO.dll, commonly found under the path 'C:\Program Files\Microsoft Office\root\Office16\'. If the DLL exists under that path in your system things should work fine. If it does not, you will have to search for it, and then browse to the DLL location to set the reference.

I have also included a procedure 'CheckDAOReference', that will check if the reference is set for you, but to run it you need to 'trust access to the VBA project object model', which I would not recommend doing unless you know what you are doing.
