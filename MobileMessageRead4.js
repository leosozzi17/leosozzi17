(function () {
    "use strict";
document.getElementById("acca3").textContent = "dentro il file js";
    var messageBanner;
    var token;

    Office.initialize = function () {
        document.getElementById("acca3").textContent = "dentro inizialize";
        
            

            getCurrentItem(); //nuovo metodo funzionante
    };
    function requestToken() {
        $.ajax({
            async: true,
            crossDomain: true,
            url: "https://login.microsoftonline.com/773689e9-a92e-49c4-9e2c-066fa07df154/oauth2/v2.0/token", //pass your tenant
            method: "POST",
            
            headers: {
                "content-type": "application/x-www-form-urlencoded",
            }
            ,
            data: {
                grant_type: "authorization_code",
                "client_id ": "39c1a9c1-fc46-4349-80fc-7f6e663e7474", //Provide your app id
                client_secret: "Adm8Q~kDcz_f~SeKzxtgJloWYcJEVyVbecZJZcE6", //Provide your client secret genereated from your app
                "scope ": "user.read%20mail.read",
            },
            success: function (response) {
                console.log(response);
                token = response.access_token;
                console.log(token);
            },
            error: function (error) {
                console.log(JSON.stringify(error));
            },
        });
    }

    function getCode() {
        var getMessageUrl = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize?';

        var xhr = new XMLHttpRequest();
        xhr.open('GET', getMessageUrl);

        xhr.onload = function (e) {
            console.log(this.response);
        }
        xhr.send();
    }

    async function getCurrentItem() {
        document.getElementById("acca3").textContent = "dentro metodo GetCurrentItem";
        const headers = new Headers();
        const bearer = 'Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6ImU0OTFyUkRiV05lM09FZm03TU1mRmFubzBLQk1xci1oWXk3OUlNWnVtZEUiLCJhbGciOiJSUzI1NiIsIng1dCI6ImpTMVhvMU9XRGpfNTJ2YndHTmd2UU8yVnpNYyIsImtpZCI6ImpTMVhvMU9XRGpfNTJ2YndHTmd2UU8yVnpNYyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC83NzM2ODllOS1hOTJlLTQ5YzQtOWUyYy0wNjZmYTA3ZGYxNTQvIiwiaWF0IjoxNjU0Njk0ODEwLCJuYmYiOjE2NTQ2OTQ4MTAsImV4cCI6MTY1NDcwMDM3MiwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkUyWmdZTWordG5ES3FSOXNoZlloSEx3VzV5cW1UWXk1OFd0N1IzUGhNK05HM2UwbStsc0IiLCJhbXIiOlsicHdkIiwicnNhIl0sImFwcF9kaXNwbGF5bmFtZSI6IkdyYXBoIEV4cGxvcmVyIiwiYXBwaWQiOiJkZThiYzhiNS1kOWY5LTQ4YjEtYThhZC1iNzQ4ZGE3MjUwNjQiLCJhcHBpZGFjciI6IjAiLCJkZXZpY2VpZCI6IjlmMjAwZDI3LWY2ZTUtNGE1ZC04ZTQ2LTg5M2Q5ZWNiNmM0OSIsImZhbWlseV9uYW1lIjoiU296emkiLCJnaXZlbl9uYW1lIjoiTGVvbmFyZG8iLCJpZHR5cCI6InVzZXIiLCJpcGFkZHIiOiIyMTMuMjYuNTIuMTk1IiwibmFtZSI6Ikxlb25hcmRvIFNvenppIiwib2lkIjoiODRmMDBkZWMtZjZiYi00N2UyLThhMDktMTVmM2U1ODkyZjMyIiwib25wcmVtX3NpZCI6IlMtMS01LTIxLTg2OTc5ODg2Ny00NDMzMDkwMzMtMTU0NDg5ODk0Mi05NDUzIiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDMyMDAxRTQ2MDU5NzQiLCJyaCI6IjAuQVVjQTZZazJkeTZweEVtZUxBWnZvSDN4VkFNQUFBQUFBQUFBd0FBQUFBQUFBQUJIQUdZLiIsInNjcCI6Ik1haWwuUmVhZCBNYWlsLlJlYWQuU2hhcmVkIE1haWwuUmVhZEJhc2ljIE1haWwuUmVhZFdyaXRlIE1haWwuUmVhZFdyaXRlLlNoYXJlZCBNYWlsLlNlbmQgTWFpbC5TZW5kLlNoYXJlZCBvcGVuaWQgcHJvZmlsZSBVc2VyLlJlYWQgZW1haWwiLCJzaWduaW5fc3RhdGUiOlsiaW5rbm93bm50d2siLCJrbXNpIl0sInN1YiI6IkN6LXFkR2N4cnNvLVV4empBNVJFeDczV0gzeTYtYkJ6d043Y3BFS01DN0UiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiRVUiLCJ0aWQiOiI3NzM2ODllOS1hOTJlLTQ5YzQtOWUyYy0wNjZmYTA3ZGYxNTQiLCJ1bmlxdWVfbmFtZSI6Imxzb3p6aUBzb2Z0ZWFtLml0IiwidXBuIjoibHNvenppQHNvZnRlYW0uaXQiLCJ1dGkiOiJwakl0aFFpMm5FMm9ZMTNpRjMwb0FBIiwidmVyIjoiMS4wIiwid2lkcyI6WyJiNzlmYmY0ZC0zZWY5LTQ2ODktODE0My03NmIxOTRlODU1MDkiXSwieG1zX3N0Ijp7InN1YiI6InpLeVo2V0NndFJzSGVuN1BHa3o4elhqS3FIUUNXYkJzcXdpSVcwVzZWMUEifSwieG1zX3RjZHQiOjE1ODE0MDU4NDJ9.iSCw3HGmzbtW5U6wSN2msck8slZWfRezbhifOSLZvExMyuL0-Ucm9SrhP8oIre6itil8HvEde24Er6BIbkdbQWT8tBJMbl8_kn7M5Jff0roPYVvwSbOIIeZu_sJBYiwhf4rv2-AnDLjLMKvT9tkPq1me2qjvc4hVs9lvVf-w5t28OJOUUdiTpLS0V8fRsoiMQTOeOz5blW5bMOOLyBf8EjvZ_pGPWWTSnc5GTIvjn0X89qT5vPlSWul-UmZN01F2fvcDlPZ28TaOCw3rbf6c8nwqJ2KDi6Ep_nCnKvHnESXh6eSDM1s1M4w0PbJVMOTj-iVHU-g2JoFm5uaF0Og8-Q';
        headers.append("Authorization", bearer);

        var x;

        var getMessageUrl = 'https://graph.microsoft.com/v1.0/me/messages/' + getItemRestId() + '/$value';

        document.getElementById("acca3").textContent = getMessageUrl;

        var xhr = new XMLHttpRequest();
        xhr.open('GET', getMessageUrl);
        xhr.setRequestHeader("Authorization", bearer);
        xhr.onload = function (e) {
            x = this.response;
            console.log(x);
            var a = document.createElement("a");
        a.href = window.URL.createObjectURL(new Blob([x], { type: 'text/plain' }));
        a.download = "demo.eml";
            a.click();
            console.log(a);
            if (a !== undefined) {
                //document.getElementById("acca3").textContent = a;
            }
        }
        xhr.send();


        //console.log(getItemRestId());


        //const credential = new DeviceCodeCredential(tenantId, clientId, clientSecret);
        //const authProvider = new TokenCredentialAuthenticationProvider(credential, {
        //    scopes: [scopes]
        //});

        //const client = Client.initWithMiddleware({
        //    debugLogging: true,
        //    authProvider
            // Use the authProvider object to create the class.
        /*});*/
        

        //const {
        //    Client
        //} = require("@microsoft/microsoft-graph-client");
        //const {
        //    TokenCredentialAuthenticationProvider
        //} = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
        //const {
        //    AuthorizationCodeCredential
        //} = require("@azure/identity");

        //const credential = new AuthorizationCodeCredential(
        //    "773689e9-a92e-49c4-9e2c-066fa07df154",
        //    "39c1a9c1-fc46-4349-80fc-7f6e663e7474",
        //    "<AUTH_CODE_FROM_QUERY_PARAMETERS>",
        //    "http://localhost/abc"
        //);
        //const authProvider = new TokenCredentialAuthenticationProvider(credential, {
        //    scopes: [scopes]
        //});
    }

    function getItemRestId() {
        if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
            // itemId is already REST-formatted.
            return Office.context.mailbox.item.itemId;
        } else {
            // Convert to an item ID for API v2.0.
            return Office.context.mailbox.convertToRestId(
                Office.context.mailbox.item.itemId,
                Office.MailboxEnums.RestVersion.v2_0
            );
        }
    }

    function base64toBlob(base64Data, contentType) {
        contentType = contentType || '';
        var sliceSize = 1024;
        var byteCharacters = atob(base64Data);
        var bytesLength = byteCharacters.length;
        var slicesCount = Math.ceil(bytesLength / sliceSize);
        var byteArrays = new Array(slicesCount);

        for (var sliceIndex = 0; sliceIndex < slicesCount; ++sliceIndex) {
            var begin = sliceIndex * sliceSize;
            var end = Math.min(begin + sliceSize, bytesLength);

            var bytes = new Array(end - begin);
            for (var offset = begin, i = 0; offset < end; ++i, ++offset) {
                bytes[i] = byteCharacters[offset].charCodeAt(0);
            }
            byteArrays[sliceIndex] = new Uint8Array(bytes);
        }
        return new Blob(byteArrays, { type: contentType });
    }

    function download(filename, text) {
        var downloadblob = base64toBlob(text);
        console.log(downloadblob);
        if (window.navigator && window.navigator.msSaveOrOpenBlob) {
            window.navigator.msSaveOrOpenBlob(downloadblob, filename);
            return;
        }
        const data = window.URL.createObjectURL(downloadblob);
        var element = document.createElement('a');
        element.setAttribute('href', data);
        element.setAttribute('download', filename);
        element.style.display = 'none';
        document.body.appendChild(element);
        element.click();
        document.body.removeChild(element);
    }
    function getSoapEnvelope(request) {
        // Wrap an Exchange Web Services request in a SOAP envelope.
        var result =

            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
            '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '  <soap:Header>' +
            '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
            '  </soap:Header>' +
            '  <soap:Body>' +

            request +

            '  </soap:Body>' +
            '</soap:Envelope>';

        return result;
    }

    function GetItem() {
        var results =
            '  <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
            '    <ItemShape>' +
            '      <t:BaseShape>IdOnly</t:BaseShape>' +
            '      <t:IncludeMimeContent>true</t:IncludeMimeContent>' +
            '      <AdditionalProperties xmlns="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '        <FieldURI FieldURI="item:Subject" />' +
            '      </AdditionalProperties>' +
            '    </ItemShape>' +
            '    <ItemIds>' +
            '      <t:ItemId Id="' + Office.context.mailbox.item.itemId + '" />' +
            '    </ItemIds>' +
            '  </GetItem>';

        return results;
    }




    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();