/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize Office UI Fabric components (dropdowns)
            $(".ms-Dropdown").Dropdown();

            // Add event handlers
            $('#add-plain-text').click(addPlainText);
            $('#add-html').click(addHtml);
            $('#add-matrix').click(addMatrix);
            $('#add-office-table').click(addOfficeTable);
            $('#add-open-xml').click(addOpenXml);
            $('#get-selected-plain-text').click(getSelectedPlainText);
            $('#get-selected-html').click(getSelectedHTML);
            $('#selected-data-dialog-more').click(hideSelectedDataDialog);
            $('#selected-data-dialog-got-it').click(hideSelectedDataDialog);
        });
    };

    // Add data (plain text) to the current document selection
    function addPlainText() {
        var text = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod' +
            'tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, ' +
            'quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ' +
            'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore ' +
            'eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt ' +
            'in culpa qui officia deserunt mollit anim id est laborum.';

        // Set selection
        Office.context.document.setSelectedDataAsync(text, { coercionType: 'text' }, onSelectionSet);
    }

    // Add data (HTML) to the current document selection
    function addHtml() {
        var elements = $('<div>')
            .append($('<h2>').text('Lorem ipsum dolor'))
            .append($('<p>').html('Duis aute irure dolor in <strong>reprehenderit in</strong>' +
                                  'voluptate velit esse cillum dolore.'));
        var html = elements.html();

        // Set selection
        Office.context.document.setSelectedDataAsync(html, { coercionType: 'html' }, onSelectionSet);
    }

    // Add data (Matrix) to the current document selection
    function addMatrix() {
        var matrix = [["Header", "Header"],
                ["Entry", "Entry"],
                ["Entry", "Entry"],
                ["Entry", "Entry"]];

        // Set selection
        Office.context.document.setSelectedDataAsync(matrix, { coercionType: 'matrix' }, onSelectionSet);
    }

    // Add data (Office Table) to the current document selection
    function addOfficeTable() {
        var table = new Office.TableData();
        table.headers = [['Header', 'Header']];
        table.rows = [['Entry', 'Entry'], ['Entry', 'Entry'], ['Entry', 'Entry']];

        // Set selection
        Office.context.document.setSelectedDataAsync(table, { coercionType: 'table' }, onSelectionSet);
    }

    // Add data (Open XML) to the current document selection
    function addOpenXml() {
        // Get the selected file
        var file = $('#ooxml-file').val();

        // Get the file contents
        $.ajax({
            url: '/../../OOXML/' + file,
            type: 'GET',
            dataType: 'text',
            success: function (data) {
                // Set selection
                Office.context.document.setSelectedDataAsync(data, { coercionType: 'ooxml' }, onSelectionSet);
            },
            error: function (e) {
                // TODO: Handle error
            }
        });
    }

    // Callback function for the asynchronous write function
    function onSelectionSet(asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            // TODO: Handle error
        }
    }

    // Get the selected data as plain text
    function getSelectedPlainText() {
        getSelectedData('text');
    }

    // Get the selected data as HTML
    function getSelectedHTML() {
        getSelectedData('html');
    }

    function getSelectedData(coercionType) {
        Office.context.document.getSelectedDataAsync(coercionType, function (asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                // TODO: Handle error
            }
            else {
                showSelectedDataDialog(asyncResult.value);
            }
        });
    }

    // Show the selected data dialog
    function showSelectedDataDialog(text) {
        if (text.length === 0) {
            text = 'No selected data was found';
        }

        // Set the dialog text
        $('#selected-data-dialog-text').text(text);
        $('#selected-data-dialog').show();
    }

    // Hide the selected data dialog
    function hideSelectedDataDialog() {
        $('#selected-data-dialog').hide();
    }
})();