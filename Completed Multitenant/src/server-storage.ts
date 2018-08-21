// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file wraps some of the functions of the node-persist library. 
*/

import * as storage from 'node-persist';

export class ServerStorage {

    public static persist(key: string, data: any) {
        storage.initSync();
        storage.setItemSync(key, data);
    }

    public static retrieve(key: string) {
        storage.initSync();
        if (storage.getItemSync(key)) {
            return storage.getItemSync(key);
        } else {
            return null;
        }
    }
    
    public static clear() {
        storage.initSync();
        storage.clearSync();
    }
}