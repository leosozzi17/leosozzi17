(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function () {
        $(document).ready(function () {

            var accessToken = "eyJ0eXAiOiJKV1QiLCJub25jZSI6InlPeDRSNWNiNlRHSVlwbDVLVW9MME9aMVZiTDhmLUt0S0psWFhQU3pvdTgiLCJhbGciOiJSUzI1NiIsIng1dCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSIsImtpZCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC83NzM2ODllOS1hOTJlLTQ5YzQtOWUyYy0wNjZmYTA3ZGYxNTQvIiwiaWF0IjoxNjU2NTcxMzYyLCJuYmYiOjE2NTY1NzEzNjIsImV4cCI6MTY1NjU3NTI2MiwiYWlvIjoiRTJaZ1lPaUo1L3RUYzNuNnZmNmZoUWtDaC84N0FnQT0iLCJhcHBfZGlzcGxheW5hbWUiOiJFeHBvcnQgTWFpbCIsImFwcGlkIjoiMzljMWE5YzEtZmM0Ni00MzQ5LTgwZmMtN2Y2ZTY2M2U3NDc0IiwiYXBwaWRhY3IiOiIxIiwiaWRwIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvNzczNjg5ZTktYTkyZS00OWM0LTllMmMtMDY2ZmEwN2RmMTU0LyIsImlkdHlwIjoiYXBwIiwib2lkIjoiY2Y4NTNhZTUtODg0Zi00ODlkLThmZmMtMGYzZTJlZTczZGVhIiwicmgiOiIwLkFVY0E2WWsyZHk2cHhFbWVMQVp2b0gzeFZBTUFBQUFBQUFBQXdBQUFBQUFBQUFCSEFBQS4iLCJzdWIiOiJjZjg1M2FlNS04ODRmLTQ4OWQtOGZmYy0wZjNlMmVlNzNkZWEiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiRVUiLCJ0aWQiOiI3NzM2ODllOS1hOTJlLTQ5YzQtOWUyYy0wNjZmYTA3ZGYxNTQiLCJ1dGkiOiI1QUFUUC1PLXhVdVFFcGIwWEx4eUFBIiwidmVyIjoiMS4wIiwid2lkcyI6WyIwOTk3YTFkMC0wZDFkLTRhY2ItYjQwOC1kNWNhNzMxMjFlOTAiXSwieG1zX3RjZHQiOjE1ODE0MDU4NDJ9.s780MH9lnpkGaVazgXQSo1XCMg3o20FFac5RCWvGrS6mStqsgnLylrTppkgrzrcRI3MSXAL4vSdxwZ-Rnkwpx5hIgADBywHg48bI5SLDOu9ZNaa50NQM9hFFnv9cUenUvFBrP83kKmYNmBa8qxsghbj_U8B8oNMDl7Gc4Y4Acn5dJ33SYNqOEAGw6Tuc7tX9f6_fW6RPrpACq4AOYpdwYY-ZRg6msrr_Hp0e3KQMYT8nsHSeenZuVanuu8aZKAJx7yotYm9J4OgdHOI5OoXSgrURfuKtUQWKYsxpZkuwlARvF3LoizennlwfjX79L1ULlkjpK1s69cc4TYrrx2UC0A";
            //requestToken();

            getCurrentItem(accessToken); //funzione per ottenere mail
        });
    };

    function requestToken() { //POST per ottenere access Token
        $.ajax({
            async: true,
            crossDomain: true,
            url: "https://login.microsoftonline.com/773689e9-a92e-49c4-9e2c-066fa07df154/oauth2/v2.0/token", //pass your tenant
            method: "POST",
            
            headers: {
                "content-type": "application/x-www-form-urlencoded",
            }
            ,
            //data: {
            //    grant_type: "client_credentials",
            //    "client_id": "9505d484-e134-468e-a740-aba39fb369b4", //Provide your app id
            //    client_secret: "PCB8Q~1qfmm8JeiGvRovZ6ALQ8ufsUAvHqy1yaBS", //Provide your client secret genereated from your app
            //    "scope": "https://graph.microsoft.com/.default" //Mail.Read
            //},
            data: {
                grant_type: "client_credentials",
                "client_id": "39c1a9c1-fc46-4349-80fc-7f6e663e7474", //Provide your app id
                client_secret: "JIu8Q~-PGU3SpzZU5D9DnXnm2uxV33eBWCb7rcI0", //Provide your client secret genereated from your app
                "scope": "api://39c1a9c1-fc46-4349-80fc-7f6e663e7474/.default"
            },
            success: function (response) {
                console.log(response);
                //token = response.access_token;
                //console.log(token);
            },
            error: function (error) {
                console.log(JSON.stringify(error));
            },
        });
    }

    async function getCurrentItem(aT) { //GET messagge [e download (da separare successivamente)]
        const headers = new Headers();
        const bearer = 'Bearer ' + aT;
        headers.append("Authorization", bearer);

        var mime;

        var getMessageUrl = 'https://graph.microsoft.com/v1.0/users/84f00dec-f6bb-47e2-8a09-15f3e5892f32/messages/' + getItemRestId() + '/$value';
        //var getMessageUrl = 'https://graph.microsoft.com/v1.0/me/messages/' + getItemRestId() + '/$value';

        var xhr = new XMLHttpRequest();
        xhr.open('GET', getMessageUrl);
        xhr.setRequestHeader("Authorization", bearer);
        xhr.onload = function (e) {
            mime = this.response;
            if (this.status === 200) {
                document.getElementById("acca3").innerHTML = "Mail esportata con successo";
                //invio mail al server ...
                //download(mime);
            }
            else {
                document.getElementById("acca3").innerHTML = "Errore durante l'esportazione della mail: " + this.statusText;
                console.log(mime);
            }
        }
        xhr.send();       
    }

    function downloadMail(mail) { //downlo email come file .eml
        var a = document.createElement("a");
        a.href = window.URL.createObjectURL(new Blob([mail], { type: 'text/plain' }));
        a.download = "demo.eml";
        a.click();
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

})();
