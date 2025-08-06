/*
https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/samples/msal-node-samples/custom-INetworkModule-and-network-tracing/HttpClientAxios.ts
 */

const axios = require("axios");
const { HttpsProxyAgent } = require("https-proxy-agent");
const HttpMethod = {
    GET: "get",
    POST: "post",
};
const proxyUrl = process.env['HTTPS_PROXY'] || process.env['HTTP_PROXY'];
if (!proxyUrl) {
    throw new Error('Missing HTTP/S_PROXY env');
}
const proxyAgent = new HttpsProxyAgent(proxyUrl);
/**
 * This class implements the API for network requests.
 */
class HttpClientAxios {
    /**
     * Http Get request
     * @param {string} url
     * @param {object} [options]
     * @returns {Promise<object>}
     */
    async sendGetRequestAsync(url, options) {
        const request = {
            method: HttpMethod.GET,
            url: url,
            headers: options && options.headers,
            validateStatus: () => true,
            httpsAgent: proxyAgent,
        };

        const response = await axios(request);
        return {
            headers: response.headers,
            body: response.data,
            status: response.status,
        };
    }

    /**
     * Http Post request
     * @param {string} url
     * @param {object} [options]
     * @param {number} [cancellationToken]
     * @returns {Promise<object>}
     */
    async sendPostRequestAsync(url, options, cancellationToken) {
        const request = {
            method: HttpMethod.POST,
            url: url,
            data: (options && options.body) || "",
            timeout: cancellationToken,
            headers: options && options.headers,
            validateStatus: () => true,
            httpsAgent: proxyAgent
        };

        const response = await axios(request);
        return {
            headers: response.headers,
            body: response.data,
            status: response.status,
        };
    }
}

module.exports = { HttpClientAxios };