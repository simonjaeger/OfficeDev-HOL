/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Add event handlers
            $('#read-data-from-selection').click(readDataFromSelection);
            $('#write-data-to-selection').click(writeDataToSelection);
            $('#create-binding').click(createBinding);
            $('#write-data-to-binding').click(writeDataToBinding);
            $('#read-data-from-binding').click(readDataFromBinding);
        });
    };

    // Read data from the current selection and log it in the JavaScript Console
    function readDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (asyncResult) {
                if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                    // TODO: Handle error
                }
                else {
                    console.log('Selection: ' + asyncResult.value);
                }
            });
    }

    // Write data to the current selection 
    function writeDataToSelection() {
        Office.context.document.setSelectedDataAsync('Hello #Office365Dev', {
            coercionType: Office.CoercionType.Text
        }, function (asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                // TODO: Handle error
            }
        });
    }

    // Create a binding named "myTable" for the "MyTable" table in the Excel sheet
    function createBinding() {
        Office.context.document.bindings.addFromNamedItemAsync('MyTable',
            Office.BindingType.Table, { id: 'myTable' }, function (asyncResult) {
                if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                    // TODO: Handle error
                }
                else {
                    // We only need to create the binding once, so let's disable
                    // this functionality at this point
                    $('#create-binding').attr('disabled', 'disabled');

                    // Enable the buttons for interacting with the binding
                    $('#write-data-to-binding').removeAttr('disabled');
                    $('#read-data-from-binding').removeAttr('disabled');
                }
            });
    }

    // Pass the "myTable" binding to a callback (assuming it has
    // already been created)
    function getBinding(callback) {
        Office.context.document.bindings.getByIdAsync('myTable',
          function (asyncResult) {
              if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                  // TODO: Handle error
              }
              else {
                  var binding = asyncResult.value;
                  callback(binding);
              }
          });
    }

    // Write data to the "myTable" binding 
    function writeDataToBinding() {
        getBinding(function (binding) {
            // Create the matrix
            var data = [
                ['Entry', 'Entry', 'Entry', 'Entry'], // Row 1
                ['Entry', 'Entry', 'Entry', 'Entry'], // Row 2
                ['Entry', 'Entry', 'Entry', 'Entry']  // Row 3
            ];

            // Write data (add rows)
            binding.addRowsAsync(data,
                function (asyncResult) {
                    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                        // TODO: Handle error
                    }
                });
        });
    }

    // Read data from the "myTable" binding and log it in the JavaScript Console
    function readDataFromBinding() {
        getBinding(function (binding) {
            // Read data
            binding.getDataAsync({ coercionType: Office.CoercionType.Table },
                function (asyncResult) {
                    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                        // TODO: Handle error
                    }
                    else {
                        // Log data
                        console.log(asyncResult.value);
                    }
                });
        });
    }
})();