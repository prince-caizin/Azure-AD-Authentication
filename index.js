const express = require('express');
const msal = require('@azure/msal-node');
const config = {
    auth: {
      clientId: '0973b7e0-879c-45d8-94aa-17af6eb97d9a',
      authority: 'https://login.microsoftonline.com/83afea21-49a2-4502-8264-e11560d9fe5a',
      clientSecret: '-na8Q~MF5tafuvQaEmendgglCUjRDjSipgmsHbUd',
    },
    cache: {
      cacheLocation: 'sessionStorage',
      storeAuthStateInCookie: false
    }
  };

const app = express();

const pca = new msal.ConfidentialClientApplication(config);

app.get('/', (req, res) => {
  const authCodeUrlParameters = {
    scopes: ['user.read'],
    redirectUri: 'http://localhost:3000/redirect'
  };

  pca.getAuthCodeUrl(authCodeUrlParameters)
    .then((response) => {
      console.log(response);
      res.redirect(response);
    })
    .catch((error) => {
      console.log(error);
      res.status(500).send(error);
    });
});

app.get('/redirect', (req, res) => {
  const tokenRequest = {
    code: req.query.code,
    scopes: ['user.read'],
    redirectUri: 'http://localhost:3000/redirect'
  };

  pca.acquireTokenByCode(tokenRequest)
    .then((response) => {
      console.log('\nResponse:\n:', response);
      res.sendStatus(200);
    })
    .catch((error) => {
      console.log(error);
      res.status(500).send(error);
    });
});

app.listen(4000, () => console.log('Server listening on port 4000!'));
