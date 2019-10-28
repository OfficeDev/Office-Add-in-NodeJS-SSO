/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file defines the routes within the authRoute router.
 */

var express = require('express');
var router = express.Router();
const defaults = require('./../configure/defaults');
var fetch = require('node-fetch');
var form = require('form-urlencoded').default;
const manifest = require('office-addin-manifest');
var ssoAppData = require('./../configure/ssoAppDataSetttings');

/* GET users listing. */
router.get('/', async function(req, res, next) {
  const authorization = req.get('Authorization');
  if (authorization == null) {
     let error = new Error('No Authorization header was found.');
     next(error);
  } 
  else {
    const [schema, jwt] = authorization.split(' ');
    const appSecret = await getSecret();
    console.log(appSecret);
    const formParams = {
      client_id: process.env.CLIENT_ID,
      client_secret: appSecret,
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

  }
});

async function getSecret() {
  const manifestInfo = await manifest.readManifestFile(defaults.manifestPath);
  const appSecret = ssoAppData.getSecretFromCredentialStore(manifestInfo.displayName);
  return appSecret;
}

module.exports = router;
