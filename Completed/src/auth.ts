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

    /**
     * Get access token for the resource either from storage or by using 
     * the current exchangeable token to get a fresh token for the resource.
     * @param jwt The JSON Web Token that the client sent to get access to this application.
     * @param scopes The scopes that need to be permitted by the new token for the backend resource.
     * @param resource (optional) The resource that needs to be accessed after accquiring the new token. 
     * Do not pass a resource, if the token service is Azure AD V2 endpoint because it infers the
     * resource from the scopes.                 
     */
    async acquireTokenOnBehalfOf(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        
        // Get a new resource token by exchange, if the current one has expired or will in the next minute
        // or doesn't exist yet (e.g., the add-in is being run for the first time on this computer). Else
        // get it from storage.
        const resourceTokenExpirationTime = ServerStorage.retrieve('ResourceTokenExpiresAt');
        if (moment().add(1, 'minute').diff(await resourceTokenExpirationTime) < 1 ) {
            return ServerStorage.retrieve('ResourceToken');
        } else if (resource) {
            return this.exchangeForToken(jwt, scopes, resource);
        } else {
            return this.exchangeForToken(jwt, scopes);
        }
    }

    /**
     * Exchange the current exchangeable token for a new token to a resource.
     * @param jwt JSON Web Token that obtained via single sign on.
     * @param scopes The scopes that need to be permitted on the new token
     * @param resource (optional) The resource that needs to be accessed after accquiring the new token. 
     * Do not pass a resource, if the token service is Azure AD V2 endpoint because it infers the
     * resource from the scopes.                 .                 
     */
    private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        try {
            // The Azure AD V2 endpoint infers the intended resource from the scopes. If 
            // a redundant resource parameter is sent to it, Azure AD V2 will return an error and not send
            // the token. So we need to ensure that we don't send one, when V2 is the token endpoint.            
            const v2Params = {
                client_id: this.clientId,
                client_secret: this.clientSecret,
                grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                assertion: jwt,
                requested_token_use: 'on_behalf_of',
                scope: scopes.join(' ')
            };

            let finalParams = {};
            if (resource) {
                let v1Params  = { resource: resource };  
                for(var key in v2Params) { v1Params[key] = v2Params[key]; }
                finalParams = v1Params;
            } else {
                finalParams = v2Params;
            }

            const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
                method: 'POST',
                body: form(finalParams),
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            });

            if (res.status !== 200) {
                const exception = await res.json();
                throw exception;                
            }

            const json = await res.json();

            // Persist the token and it's expiration time.
            const resourceToken = json['access_token'];
            ServerStorage.persist('ResourceToken', resourceToken);
            const expiresIn = json['expires_in'];  // seconds until token expires.
            const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
            ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);

            return resourceToken;
        }
        catch (exception) {
            throw new UnauthorizedError('Unable to obtain an access token to the resource. ' 
                                        + JSON.stringify(exception), 
                                        exception);
        }
    }
}