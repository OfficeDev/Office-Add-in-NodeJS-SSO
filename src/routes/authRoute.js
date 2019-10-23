/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file defines the routes within the authRoute router.
 */

var express = require('express');
// var jsonwebtoken = require('jsonwebtoken'); //TODO: validate bootstrap token
var router = express.Router();
var fetch = require('node-fetch');
var form = require('form-urlencoded').default;


/* GET users listing. */
router.get('/', async function(req, res, next) {
  const authorization = req.get('Authorization');
  if (authorization == null) {
      throw new Error('No Authorization header was found.');
  }
  const [schema, jwt] = authorization.split(' ');

  // TODO: validate bootstrap token
  //const decoded = jsonwebtoken.decode(jwt, { complete: true });
  const formParams = {
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
    assertion: jwt,
    requested_token_use: 'on_behalf_of',
    scope: ['Files.Read.All'].join(' ')
  };

  const stsDomain = 'https://login.microsoftonline.com';
  const tenant = 'common';
  const tokenURLSegment = 'oauth2/v2.0/token';

  const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
    method: 'POST',
    body: form(formParams),
    headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/x-www-form-urlencoded'
    }
  });
  const json = await tokenResponse.json();
  
  res.send(json);
});

module.exports = router;
