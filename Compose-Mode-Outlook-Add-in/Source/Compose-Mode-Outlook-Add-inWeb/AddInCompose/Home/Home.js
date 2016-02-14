/// <reference path="../App.js" />

(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Add event handlers
            $('#set-subject').click(setSubject);
            $('#set-recipients').click(setRecipients);
            $('#set-body').click(setBody);
        });
    };

    // Set the item subject
    function setSubject() {
        var _item = Office.context.mailbox.item;
        var subject = _item.subject;
        subject.setAsync('Hello #Office365Dev', onDataSet)
    }

    // Set the item recipients
    function setRecipients() {
        var _item = Office.context.mailbox.item;
        var _userProfile = Office.context.mailbox.userProfile;
        var user = {
            displayName: _userProfile.displayName,
            emailAddress: _userProfile.emailAddress
        };

        // Check if the item is a message or appointment
        // in order to determine which properties to set
        if (_item.itemType === Office.MailboxEnums.ItemType.Message) {
            _item.to.setAsync([user], onDataSet);
            _item.cc.setAsync([user], onDataSet);
            _item.bcc.setAsync([user], onDataSet);
        }
        else if (_item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            _item.requiredAttendees.setAsync([user], onDataSet);
        }
    }

    // Set the item body
    function setBody() {
        var _item = Office.context.mailbox.item;
        var body = _item.body;
        var text = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod' +
            'tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, ' +
            'quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ' +
            'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore ' +
            'eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt ' +
            'in culpa qui officia deserunt mollit anim id est laborum.';
        body.setAsync(text, onDataSet)
    }

    // Callback function for the asynchronous write functions
    function onDataSet(asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            // TODO: Handle error
        }
    }
})();