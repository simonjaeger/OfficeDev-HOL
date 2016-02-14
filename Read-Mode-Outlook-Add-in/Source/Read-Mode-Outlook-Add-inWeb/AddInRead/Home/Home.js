/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Add event handlers
            $('#get-subject').click(getSubject);
            $('#get-sender').click(getSender);
            $('#get-body').click(getBody);
            $('#data-dialog-more').click(hideDataDialog);
            $('#data-dialog-got-it').click(hideDataDialog);
        });
    };

    // Get the item subject and display it
    function getSubject() {
        var _item = Office.context.mailbox.item;
        var subject = _item.subject;

        // Show data
        showDataDialog('Subject', subject);
    }

    // Get the item sender and display it
    function getSender() {
        var _item = Office.context.mailbox.item;
        var sender;

        // Check if the item is a message or appointment
        // in order to determine which property that contains
        // the sender
        if (_item.itemType === Office.MailboxEnums.ItemType.Message) {
            sender = _item.from;
        }
        else if (_item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            sender = _item.organizer;
        }

        // Show data
        showDataDialog('Sender', sender.displayName);
    }

    // Get the body of the item and display it as 
    // plain text
    function getBody() {
        var _item = Office.context.mailbox.item;
        var body = _item.body;

        // Get the body asynchronous as text
        body.getAsync(Office.CoercionType.Text, function (asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                // TODO: Handle error
            }
            else {
                // Show data
                showDataDialog('Body', asyncResult.value.trim());
            }
        });
    }

    // Show the data dialog
    function showDataDialog(title, text) {
        // Set the dialog title and text
        $('#data-dialog-title').text(title);
        $('#data-dialog-text').text(text);
        $('#data-dialog').show();
    }

    // Hide the data dialog
    function hideDataDialog() {
        $('#data-dialog').hide();
    }
})();