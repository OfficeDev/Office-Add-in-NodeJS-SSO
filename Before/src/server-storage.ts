// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file wraps some of the functions of the node-persist library. 
*/

import * as storage from 'node-persist';

export class ServerStorage {

    public static async persist(key: string, data: any) {
        await storage.init();
        await storage.setItem(key, data);
        console.log(key);
        console.log(data);
    }

    public static async retrieve(key: string) {
        await storage.init();
        if (await storage.getItem(key)) {
            return await storage.getItem(key);
        } else {
            return null;
        }
    }
    
    public static async clear() {
        await storage.init();
        await storage.clear();
    }
}