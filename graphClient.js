/*
https://learn.microsoft.com/ja-jp/graph/tutorials/node?tutorial-step=3
*/

const graph = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
const { ProxyAgent } = require('undici');
const dispatcher = new ProxyAgent(process.env.HTTP_PROXY);
const fetchOptions = {
  dispatcher
};

const getAuthenticatedClient = (msalClient, scopes) => {
  if (!msalClient) {
    throw new Error('Invalid MSAL state. Client: missing');
  }
  if (!scopes || !Array.isArray(scopes) || scopes.length === 0) {
    throw new Error('Scopes must be a non-empty array.');
  }

  // getAccessTokenメソッドを持つオブジェクトをauthProviderに渡す
  const authProvider = {
    getAccessToken: async () => {
      const response = await msalClient.acquireTokenByClientCredential({
        scopes: scopes,
      });
      return response.accessToken;
    }
  };

  return graph.Client.initWithMiddleware({
    authProvider,
    fetchOptions
  });
};

const getUsersDetails = async (msalClient, scopes) => {
  const client = getAuthenticatedClient(msalClient, scopes);
  return client
    .api('/users')
    .select('displayName,mail,userPrincipalName')
    .get();
};

// クライアントクレデンシャルフローでは /me エンドポイントは利用できません
const getUserDetails = async () => {
  throw new Error('getUserDetails is not supported in client credential flow.');
};

module.exports = {
  getUsersDetails,
  getUserDetails,
};