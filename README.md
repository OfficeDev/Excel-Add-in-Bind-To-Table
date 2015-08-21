# Excel-Add-in-Bind-To-Table
 Demonstrates how to use the JavaScript API for Office to bind an add-in to a named table in Microsoft Excel 2013 or Microsoft Word 2013, extract data from the table, react to events in the table, and set data back into the table.

**Description**

This add-in code sample binds a Stock Ticker add-in to a named table that is included in the solution project. The code that uses the Binding object of the JavaScript API for Office is included in the UpdateTable.js file of the solution. The code sample binds to a table named "Stocks" on the worksheet labeled "Sheet1."

The sample demonstrates how to establish a binding between a region in an Office file —a named table, in this case— and an add-in for Office. Once the binding has been established, the add-in adds a handler to the BindingDataChanged event of the binding. When the event handler executes, the add-in retrieves a selection of the data from the table. It then removes the event handler from the binding, updates the table, and then adds the event handler back to the binding.

For more information about the JavaScript API for Office and working with bindings, see  Binding to regions in a document or spreadsheet.

**Prerequisites**

This sample requires the following:

* Word 2013 or Excel 2013
* Visual Studio 2012 and Office Developer Tools for Visual Studio 2012
* Internet Explorer 9 or Internet Explorer 10.

**Key components**

The Stock Ticker sample add-in app contains the following:

* CodeSample_BindingAppToXLTable project.
* CodeSample_BindingAppToXLTable.xml manifest file
* Stocks.xlsx file that contains a table named "Stocks" on a spreadsheet named "Sheet1."
* CodeSample_BindingAppToXLTableWeb project
* Home.html file, which contains the HTML control for the add-in's user interface.
* Home.js file, which contains the event handler for the Office.initialize event of the add-in.
* UpdateTable.js file, which contains the self-executing anonymous function that creates the binding, adds an event handler to the binding, and contains all of the methods for getting and setting data from the binding.

**Configure the sample**

To configure the Stock Ticker, set the StartAction -> Start Document property of the CodeSample_BindingAppToTable project to 'Stocks.xslx'.

**Build the sample**

Choose the F5 key to build and deploy the add-in.

**Run and test the sample**

1. Choose the F5 key to build and deploy the add-in.
2. Insert the add-in into the Stocks.xlsx file when you debug the sample (Insert tab, Add-ins for Office button).

 ***Note***

 If you save the workbook while debugging, the add-in is persisted in the workbook. If you do so, you won't have to reinsert the  add-in into the workbook during future debugging sessions.

3. In the add-in, choose Set binding.
4. In the table on Book1.xlsx, make a change to one of the values in the first column in the right.

**Troubleshooting**

If the app fails to install, ensure that the XML in your AppManifest.xml file parses correctly.

If you change the code in the StockTicker.getStockQuotes method to call an external stock quote service, be aware that cross-domain scripting restrictions still apply.

If the app generates errors whenever you try to update the table, ensure that you have entered correct values for the tableName and bindingName variables in the UpdateTable.js file.

**Change log**

* First release. March 8, 2013.
* GitHub release. August 14, 2015.

**Related content**

* [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Binding to regions in a document or spreadsheet](http://msdn.microsoft.com/en-us/library/office/apps/fp123511.aspx)
* [Bindings object](http://msdn.microsoft.com/en-us/library/office/apps/fp160966.aspx)
* [Binding object](http://msdn.microsoft.com/en-us/library/office/apps/fp161045.aspx)
