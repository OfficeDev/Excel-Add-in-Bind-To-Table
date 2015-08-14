var StockTicker = (function () {

    // Change the "Sheet1!Stocks" and "Stocks" to match the name
    // of the table and binding that you want to use.
    var tableName = "Sheet1!Stocks",
        bindingName = "Stocks",
        binding;

    // Create the binding to the table on the spreadsheet.
    function initializeBinding() {
        Office.context.document.bindings.addFromNamedItemAsync(
          tableName,
          Office.BindingType.Table,
          { id: bindingName },
          function (results) {

              // Capture a reference to the binding and then add
              // an event handler to the binding.
              binding = results.value;
              addBindingsHandler(function () { refreshData(); });
          });
    }

    // Event handler to refresh the table when the
    // data in the table changes.
    var onBindingDataChanged = function (result) {
        refreshData();
    }

    // Add the handler to the BindingDataChanged event of the binding.
    // When the event is raised, the callback function is called.
    function addBindingsHandler(callback) {
        Office.select("bindings#" + bindingName).addHandlerAsync(
          Office.EventType.BindingDataChanged,
          onBindingDataChanged,
          function () {
              if (callback) { callback(); }
          });
    }

    // Refresh the data displayed in the bound table of the workbook. 
    // This function begins a chain of asynchronous calls that 
    // updates the bound table.
    function refreshData() {
        getBindingData();
    }

    // Get the stock symbols from the bound table and
    // then call the stock quote information service.
    function getBindingData() {
        binding.getDataAsync(
          {
              startRow: 0,
              startColumn: 0,
              columnCount: 1
          },
          function (results) {
              var bindingData = results.value,
                  stockSymbols = [];

              for (var i = 0; i < bindingData.rows.length; i++) {
                  stockSymbols.push(bindingData.rows[i][0]);
              }

              // Call the Web service or proxy to get the 
              // stock quote information.
              getStockQuotes(stockSymbols);
          });
    }

    // Call a faked 'Web service' to get new stock quotes.
    // **If you are calling a third-party Web service, you should follow
    // best practices to avoid problems with cross-domain scripting.**
    function getStockQuotes(stockSymbols) {

        var stockValues = [];

        for (var i = 0; i < stockSymbols.length; i++) {
            // Generate some random numbers and return
            // an array of arrays with stock symbols and values.
            stockValues.push([stockSymbols[i], (100 * Math.random()).toFixed(2)]);
        }

        removeHandler(function () {
            updateTable(stockValues);
        });
    }

    // Disables the BindingDataChanged event handler
    // while the table is being updated.
    function removeHandler(callback) {

        binding.removeHandlerAsync(
          Office.EventType.BindingDataChanged,
          { handler: onBindingDataChanged },
          function (results) {
              if (results.status == Office.AsyncResultStatus.Succeeded) {
                  if (callback) { callback(); }
              }
          });
    }

    // Update the TableData object referenced by the binding
    // and then update the data in the table on the worksheet.
    function updateTable(stockValues) {

        // Insert the stock quote values into a new TableData object
        var stockData = new Office.TableData(),
            newValues = [];

        for (var i = 0; i < stockValues.length; i++) {
            var stockSymbol = stockValues[i],
                newValue = [stockSymbol[1]];

            newValues.push(newValue);
        }

        stockData.rows = newValues;

        // Set the data into the third column of the table.
        binding.setDataAsync(
          stockData,
          {
              coercionType: Office.CoercionType.Table,
              startColumn: 3,
              startRow: 0

          },
          function (results) {

              // If the method succeeds, add the bindings handler
              // back to the table.
              if (results.status == Office.AsyncResultStatus.Succeeded) {
                  addBindingsHandler();
              }
          });
    }

    // Expose the initializeBinding method.
    return {
        initializeBinding: initializeBinding
    };

})();
