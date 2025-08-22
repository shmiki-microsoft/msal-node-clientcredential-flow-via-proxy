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
const fs = require('fs');
const path = require('path');
const DELTA_LINK_FILE = path.join(__dirname, 'deltaLink.txt');

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

const getUsersDelta = async (msalClient, scopes) => {
  const client = getAuthenticatedClient(msalClient, scopes);
  const users = [];
  let deltaLink = null;
  let response;

  // deltaLinkファイルがあればそれを使う
  if (fs.existsSync(DELTA_LINK_FILE)) {
    const savedDeltaLink = fs.readFileSync(DELTA_LINK_FILE, 'utf8');
    if (savedDeltaLink) {
      response = await client.api(savedDeltaLink).get();
    }
  }

  // なければ初回取得
  if (!response) {
    response = await client
      .api('/users/delta')
      .select('displayName,mail,userPrincipalName')
      .top(1)
      .get();
  }

  // PageIterator で全ページ取得
  const pageIterator = new graph.PageIterator(
    client,
    response,
    (user) => {
      users.push(user);
      return true; // 続行
    }
  );
  await pageIterator.iterate();

  // deltaLink取得
  if (response['@odata.deltaLink']) {
    deltaLink = response['@odata.deltaLink'];
  } else if (response['@odata.nextLink']) {
    let nextLink = response['@odata.nextLink'];
    let lastResponse = response;
    while (nextLink) {
      lastResponse = await client.api(nextLink).get();
      if (lastResponse.value) {
        lastResponse.value.forEach(user => users.push(user));
      }
      nextLink = lastResponse['@odata.nextLink'];
      if (lastResponse['@odata.deltaLink']) {
        deltaLink = lastResponse['@odata.deltaLink'];
        break;
      }
    }
  }

  // deltaLinkをファイルに保存
  if (deltaLink) {
    fs.writeFileSync(DELTA_LINK_FILE, deltaLink, 'utf8');
  }

  return { value: users, deltaLink };
};


module.exports = {
  getUsersDetails,
  getUserDetails,
  getUsersDelta,
};