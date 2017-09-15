// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file provides the provides server startup, authorization context creation, and the Web APIs of the add-in. 
*/

import * as fs from 'fs';
import * as https from 'https';
import * as path from 'path';
import * as express from 'express';
import * as bodyParser from 'body-parser';
import * as cors from 'cors';
import * as morgan from 'morgan';
import { AuthModule } from './auth';
import { MSGraphHelper} from './msgraph-helper';

/* Set the environment to development if not set */
const env = process.env.NODE_ENV || 'development';

/* Instantiate AuthModule to assist with JWT parsing and verification, and token acquisition. */
const auth = new AuthModule(
    /* These values are required for our application to exhcange and get access to the resource data */
    /* client_id */ '{client GUID}',
    /* client_secret */ '{client secret}',

    /* This information tells our server where to download the signing keys to validate the JWT that we received,
     * and where to get tokens. This is not configured for multi tenant; i.e., it is assumed that the source of the JWT and our application live
     * on the same tenant.
     */
    /* tenant */ 'common',
    /* stsDomain */ 'https://login.microsoftonline.com',
    /* discoveryURLsegment */ '.well-known/openid-configuration',
    /* tokenURLsegment */ '/oauth2/v2.0/token',

    /* Token is validated against the following values: */
    // Audience is the same as the client ID because, relative to the Office host, the add-in is the "resource".
    /* audience */ '{audience GUID}', 
    /* scopes */ ['access_as_user'],
    /* issuer */ 'https://login.microsoftonline.com/{O365 tenant GUID}/v2.0',
);

/* A promisified express handler to catch errors easily */
const handler =
    (callback: (req: express.Request, res: express.Response, next?: express.NextFunction) => Promise<any>) =>
        (req, res, next) => callback(req, res, next)
            .catch(error => {
                /* If the headers are already sent then resort to the built in error handler */
                if (res.headersSent) {
                    return next(error);
                }

                /**
                 * If running development environment we send the error details back.
                 * Else we send the right code and message.
                 */
                if (env === 'development') {
                    return res.status(error.code || 500).json({ error });
                }
                else {
                    return res.status(error.code || 500).send(error.message);
                }
            });

/* Create the express app and add the required middleware */
const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cors());
app.use(morgan('dev'));
app.use(express.static('public'));

/**
 * If running on development env, then use the locally available certificates.
 */
if (env === 'development') {
    const cert = {
        key: fs.readFileSync(path.resolve('./dist/certs/server.key')),
        cert: fs.readFileSync(path.resolve('./dist/certs/server.crt'))
    };
    https.createServer(cert, app).listen(3000, () => console.log('Server running on 3000'));
}
else {
    /**
     * We don't use https as we are assuming the production environment would be on Azure.
     * Here IIS_NODE will handle https requests and pass them along to the node http module
     */
    app.listen(process.env.port || 1337, () => console.log(`Server listening on port ${process.env.port}`));
}

/**
 * HTTP GET: /api/onedriveitems
 * When passed a JWT token in the header, it extracts it and
 * and exchanges for a token that has permissions to graph.
 */
app.get('/api/onedriveitems', handler(async (req, res) => {
    /**
     * Only initializes the auth the first time
     * and uses the downloaded keys information subsequently.
     */
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' });

    let graphToken = null;

    // We don't pass a resource parameter becuase the token endpoint is Azure AD V2.
    const tokenAcquisitionResponse = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);

    // If the response is a string containing 'capolids" then this is a claims
    // message from AAD that multi-factor auth is required. Pass the message to 
    // the client, which uses it to start a second sign-on. The string tells AAD
    // what additional auth factor(s) it should prompt the user to provide.
    if (tokenAcquisitionResponse.includes('capolids')) {
        const claims: string[] = [];
        claims.push(tokenAcquisitionResponse);
        return res.json(claims);
    } else {
        // The response is the token to Graph itself. Rename it so remaining code
        // is self-documenting.
        graphToken = tokenAcquisitionResponse;
    }

    // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
    // and only the top 3 folder or file names.
    const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");   

    // Graph data includes OData metadata and eTags that we don't need. 
    // Send only what is actually needed to the client: the item names.
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);   
}));

/**
 * HTTP GET: /index.html
 * Loads the add-in home page.
 */
app.get('/index.html', handler(async (req, res) => {
    return res.sendfile('index.html');
}));
