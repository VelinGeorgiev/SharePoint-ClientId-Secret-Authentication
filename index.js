const fetch = require('isomorphic-unfetch');

const tenantId = 'velingeorgiev.onmicrosoft.com'; // or also called realm
const sharePointDomain = 'velingeorgiev.sharepoint.com';
const sharePointSiteUrl = 'https://velingeorgiev.sharepoint.com/sites/aad';
const clientId = 'a4c8fa6a-9bda-4363-87d0-9b7cf7a9290d'; // <YOUR_APP_CLIENT_ID>
const secret = 'l+BmH7dIpHgwBUUsXDB7qHyeW1Cgge+kkB4j5BsvIfM='; // <YOUR_APP_SECRET>
let accessToken = '';

fetch(`https://accounts.accesscontrol.windows.net/${tenantId}/tokens/OAuth/2`, {
    method: 'post',
    body: `
    grant_type=client_credentials
    &client_id=${clientId}@${tenantId}
    &client_secret=${encodeURIComponent(secret)}
    &resource=00000003-0000-0ff1-ce00-000000000000/${sharePointDomain}@${tenantId}
    &scope=00000003-0000-0ff1-ce00-000000000000/${sharePointDomain}@${tenantId}
    `, // URLSearchParams
    headers: {
        'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8',
        'Accept': 'application/json'
    },
})
.then((res) => res.json())
.then((authResponse) => {
    console.log(`Auth SUCCESS: ${JSON.stringify(authResponse, null, 2)}`);

    accessToken = authResponse.access_token;

    // Many endpoints do not require the X-RequestDigest in the header, but some do.
    return fetch(`${sharePointSiteUrl}/_api/contextinfo`, {
        method: 'post',
        headers: {
            'Accept': 'application/json',
            'authorization': `Bearer ${accessToken}`
        }
    });
})
.then((res) => res.json())
.then((contextInfoResponse) => {
    console.log(`ContextInfo SUCCESS: ${JSON.stringify(contextInfoResponse, null, 2)}`);

    return fetch(`${sharePointSiteUrl}/_api/web`, {
        headers: {
            'Accept': 'application/json',
            'authorization': `Bearer ${accessToken}`,
            "X-RequestDigest": contextInfoResponse.FormDigestValue
        }
    });
})
.then((res) => res.json())
.then((webResponse) => {

    console.log(`Get web SUCCESS: ${JSON.stringify(webResponse, null, 2)}`);
})
.catch((err) => {
    console.error('Auth ERROR:', err);
});