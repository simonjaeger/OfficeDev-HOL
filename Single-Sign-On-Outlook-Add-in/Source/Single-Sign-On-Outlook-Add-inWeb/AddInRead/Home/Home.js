/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize Office UI Fabric components (spinner)
            if (typeof fabric !== "") {
                if ('Spinner' in fabric) {
                    var element = document.querySelector('.ms-Spinner');
                    if (element) {
                        var component = new fabric['Spinner'](element);
                    }
                }
            }

            // Initialize input fields
            $('#username').val('#Office365Dev');
            $('#password').val('#Office365Dev');

            // Add event handlers
            $('#sign-in').click(signIn);
            $('#dialog-ok').click(hideDialog);

            // Authenticate silently (without credentials)
            authenticate(null, function (response) {
                // Hide spinner and show button
                $('#sign-in').show();
                $('#spinner').hide();

                if (!response) {
                    // Enable sign-in
                    $('#sign-in').removeAttr('disabled');
                    $('#username').removeAttr('disabled');
                    $('#password').removeAttr('disabled');
                }
                else {
                    // Display user data
                    showDialog('Hi developer!', 'Office 365 user ' + '(' +
                        Office.context.mailbox.userProfile.emailAddress + ') ' +
                        'has automatically signed in as user: ' +
                        response.displayName + ' (' +
                        response.credentials.username + ').');
                }
            });
        });
    };

    // Try to authenticate, if credentials for a user is provided - we
    // will try to map it with the Office 365 user in the backend
    function authenticate(credentials, callback) {
        Office.context.mailbox.getUserIdentityTokenAsync(function (asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                // TODO: Handle error
                callback();
            }
            else {
                var token = asyncResult.value;

                // POST the credentials to the Web API
                $.ajax({
                    type: 'POST',
                    url: '../../api/sso',
                    data: JSON.stringify({
                        identityToken: token,
                        hostUri: window.location.href.split('?')[0],
                        credentials: credentials
                    }),
                    contentType: 'application/json;charset=utf-8'
                }).done(function (response) {
                    // TODO: Validate response
                    callback(response);
                }).fail(function (error) {
                    // TODO: Handle error
                    callback();
                });
            }
        });
    }

    // Sign in using the provided credentials
    function signIn() {
        // Show spinner and hide button
        $('#sign-in').hide();
        $('#spinner').show();

        // Disable input fields
        $('#username').attr('disabled', 'disabled');
        $('#password').attr('disabled', 'disabled');

        // Get credentials
        var credentials = {
            username: $('#username').val(),
            password: $('#password').val(),
        }

        // Authenticate
        authenticate(credentials, function (response) {
            if (!response) {
                // Hide spinner and show button
                $('#sign-in').show();
                $('#spinner').hide();

                // Enable input fields
                $('#username').removeAttr('disabled');
                $('#password').removeAttr('disabled');

                // Display error
                showDialog('Oops!', 'Something happened... make sure that ' +
                    'the credentials are valid (for an entry in the user service).');
            }
            else {
                // Reload page and try SSO
                location.reload();
            }
        });
    }

    // Show the dialog
    function showDialog(title, text) {
        // Set the dialog title and text
        $('#dialog-title').text(title);
        $('#dialog-text').text(text);
        $('#dialog').show();
    }

    // Hide the dialog
    function hideDialog() {
        $('#dialog').hide();
    }
})();