// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file provides the provides authorization context, token verification, and token acquisition. 
*/

import 'isomorphic-fetch';
import { Request } from 'express';
import * as jsonwebtoken from 'jsonwebtoken';
import form from 'form-urlencoded';
import * as moment from 'moment';
import { ServerStorage } from './server-storage';
import { ServerError, UnauthorizedError } from './errors';

export class AuthModule {
    keys: { [kid: string]: string };
    isInitialized: boolean;

    /**
     * Initializes the AuthHelper
     * @param clientId The registration ID of the application that needs access to a resource.
     * @param clientSecret The registration secret of the application that needs access to a resource.
     * @param tenant The tenant where the user accounts are to be looked up.
     * For microsoft.com accounts or for MSA accounts its "common".
     * @param stsDomain The domain of the secure token service (STS).
     * @param discoveryURLsegment The relative URL where the STS provides token signing keys.
     * @param tokenURLsegment The relative URL where the STS provides tokens.
     * @param audience The audience to whom the access token is given; that is, the resource.
     * @param scopes The permissions the application needs to the resource.
     * @param issuer The issuer that provided the token itself.
     */
    constructor(
        public clientId: string,
        public clientSecret: string,
        public tenant: string,
        public stsDomain: string,
        public discoveryURLsegment: string,
        public tokenURLsegment: string,
        public audience: string,
        public scopes: string[],
        public issuer: string
    ) { 
    }

    /**
     * Download the signing keys and store them for reuse
     * @param force Force the re-initialzation of the helper.
     */
    async initialize(force: boolean = false) {
        if (this.keys == null || force) {
            this.isInitialized = false;
            this.keys = await this.downloadSigningKeys();
            this.isInitialized = true;
        }
    }

    /**
     * Download the tenant's well known open id configuration.
     * Extract the jwks_uri for the JWT signing keys.
     * Download the signing keys and store them in memory as a key value pair
     * of kid and key. Also enclose the key in a BEGIN CERTIFICATE and
     * END CERTIFICATE tag.
     * (https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-token-and-claims#validating-tokens)
     */
    private async downloadSigningKeys() {
        try {
            const urlRes = await fetch(`${this.stsDomain}/${this.tenant}/${this.discoveryURLsegment}`, {
                headers: {
                    'Content-Type': 'application/json'
                }
            });

            if (urlRes.status !== 200) {
                let error = await urlRes.text();
                throw new ServerError('Unable to download openid-configuration.', urlRes.status, error);
            }

            const { jwks_uri } = await urlRes.json();
            const res = await fetch(jwks_uri);
            const json = await res.json();
            const { keys } = json;
            let signing_keys: { [kid: string]: string } = {};
            for (const key of keys) {
                const { kid, x5c } = key;
                const [signing_key] = x5c;
                signing_keys[kid] = `-----BEGIN CERTIFICATE-----
${signing_key}
-----END CERTIFICATE-----`;
            };
            return signing_keys;
        }
        catch (exception) {
            throw new ServerError('Unable to download JWT signing keys.', 500, exception);
        }
    };

    /**
     * Verify the JWT with the appropirate signing key.
     * Upon successful validation, return the payload.
     * @param req express request parameter
     */
    verifyJWT(req: Request, assertions?: { [field: string]: string }) {
        try {
            const authorization = req.get('Authorization');
            if (authorization == null) {
                throw new UnauthorizedError('No Authorization header was found.');
            }

            const [schema, jwt] = authorization.split(' ');
            const decoded = jsonwebtoken.decode(jwt, { complete: true });
            
            /* Check return decoded type is as expected */
            if (!((<{[key:string] :any;}>decoded).header !== undefined)) throw new UnauthorizedError('Unable to verify JWT');
           
            const header = (<{[key:string] :any;}>decoded).header;
            const payload = (<{[key:string] :any;}>decoded).payload;

            /* Ensure other parameters of the payload are consistent. */
            for (const assertion of Object.keys(assertions)) {
                if (payload[assertion] !== assertions[assertion]) {
                    throw new UnauthorizedError(`JWT assertion failed: ${assertion} was ${payload[assertion]}; expected ${assertions[assertion]}`);
                }
            }

            if (schema !== 'Bearer') {
                throw new UnauthorizedError('Malformed Authorization header.');
            }

            jsonwebtoken.verify(jwt, this.keys[header.kid], { audience: this.audience, issuer: this.issuer });
            return { user: payload, jwt };
        }
        catch (exception) {
            throw new UnauthorizedError('Unable to verify JWT.' + exception.message, exception);
        }
    }
    
}
