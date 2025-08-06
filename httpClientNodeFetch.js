/*
https://github.com/AzureAD/microsoft-authentication-library-for-js/issues/6527#issuecomment-2073238927
 */

const { HttpsProxyAgent } = require("https-proxy-agent");
const fetch = require('node-fetch');

const proxyUrl = process.env['HTTPS_PROXY'] || process.env['HTTP_PROXY'];
if (!proxyUrl) {
    throw new Error('Missing HTTP/S_PROXY env');
}

const proxyAgent = new HttpsProxyAgent(proxyUrl);

class HttpClientNodeFetch {
    sendGetRequestAsync(url, options) {
        return this.sendRequestAsync(url, 'GET', options);
    }
    sendPostRequestAsync(url, options) {
        return this.sendRequestAsync(url, 'POST', options);
    }

    async sendRequestAsync(
        url,
        method,
        options = {},
    ) {
        try {
            const requestOptions = {
                method: method,
                headers: options.headers,
                body: method === 'POST' ? options.body : undefined,
                agent: proxyAgent,
            };

            console.log('>>> url', url, requestOptions);

            const response = await fetch(url, requestOptions);
            const data = await response.json();

            const headersObj = {};
            response.headers.forEach((value, key) => {
                headersObj[key] = value;
            });

            return {
                headers: headersObj,
                body: data,
                status: response.status,
            };
        } catch (err) {
            console.error('CustomRequest', err);
            throw new Error('Custom request error');
        }
    }
}

module.exports = { HttpClientNodeFetch };