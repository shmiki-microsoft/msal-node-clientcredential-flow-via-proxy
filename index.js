const msal = require('@azure/msal-node');
const { HttpClientNodeFetch } = require('./httpClientNodeFetch.js');
// const { HttpClientAxios } = require('./HttpClientAxios.js');
const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: process.env.CLOUD_INSTANCE + process.env.TENANT_ID,
        clientSecret: process.env.CLIENT_SECRET,
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: 'verbose',
        },
        networkClient: new HttpClientNodeFetch(),
        // networkClient: new HttpClientAxios(),
        // networkClient: new HttpClient(process.env.HTTP_PROXY),
    },
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

const tokenRequest = {
    scopes: ["https://graph.microsoft.com/.default"],
    skipCache: false, // false:use cache, true: use no cache
};

// Microsoft Entra ID と認証の上トークンを取得する
cca.acquireTokenByClientCredential(tokenRequest).then((response) => {
    console.log("acquireTokenByClientCredential call 1st time");
    console.log(JSON.stringify(response));

    //MSAL Node により自動的にメモリキャッシュされたトークンを取る
    cca.acquireTokenByClientCredential(tokenRequest).then((response) => {
        console.log("acquireTokenByClientCredential call 2nd time");
        console.log(JSON.stringify(response));
        }).catch((error) => {
            console.log(JSON.stringify(error));
        });

}).catch((error) => {
    console.log(JSON.stringify(error));
});

// 非同期処理のため、メモリキャッシュされる前に動作する。そのため Microsoft Entra ID と認証の上トークンを取得する
cca.acquireTokenByClientCredential(tokenRequest).then((response) => {
    console.log("acquireTokenByClientCredential call 4th time");
    console.log(JSON.stringify(response));
    }).catch((error) => {
        console.log(JSON.stringify(error));
    });

