//@ts-check
/*
   Copyright 2019 Smart-Tech Controle e Automação

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

/**
 * Compares values for equality, includes special handling for arrays. Fixes #33
 * @param {number|string|Array} a
 * @param {number|string|Array} b 
 */
function equals(a, b) {
    if (a === b) return true;
    if (a == null || b == null) return false;
    if (Array.isArray(a) && Array.isArray(b)) {
        if (a.length != b.length) return false;

        for (var i = 0; i < a.length; ++i) {
            if (a[i] !== b[i]) return false;
        }
        return true;
    }
    return false;
}

const MIN_UPDATE_RATE = 100;

module.exports = function (RED) {

    const EventEmitter = require('events').EventEmitter;
    const opcda = require('opc-da');
    const { OPCGroupStateManager, OPCItemManager, OPCSyncIO } = opcda;
    const { ComServer, Session, Clsid } = opcda.dcom;

    function generateStatus(status, val) {
        let obj;

        if (typeof val != 'string' && typeof val != 'number' && typeof val != 'boolean') {
            val = RED._("opc-da.status.online");
        }

        switch (status) {
            case 'online':
                obj = { fill: 'green', shape: 'dot', text: val.toString() };
                break;
            case 'badvalues':
                obj = { fill: 'yellow', shape: 'dot', text: RED._("opc-da.status.badvalues") };
                break;
            case 'offline':
                obj = { fill: 'red', shape: 'dot', text: RED._("opc-da.status.offline") };
                break;
            case 'connecting':
                obj = { fill: 'yellow', shape: 'dot', text: RED._("opc-da.status.connecting") };
                break;
            default:
                obj = { fill: 'grey', shape: 'dot', text: RED._("opc-da.status.unknown") };
        }
        return obj;
    }

    RED.httpAdmin.get('/opc-da/browseItems', RED.auth.needsPermission('opc-da.list'), function (req, res) {
        let params = req.query

        async function brosweItems() {
            let self = this;
            let session = new Session();
            session = session.createSession(params.domain, params.username, params.password);
            session.setGlobalSocketTimeout(params.timeout);

            let comServer = new ComServer(new Clsid(params.clsid), params.address, session);
            
            comServer.on("disconnected", function(){
                throw new Error("Disconnected from the server.");
            });
            comServer.on("e_classnotreg", function(){
                throw new Error("The given Clsid is not registered on the server.");
            });

            await comServer.init();
            
            let comObject = await comServer.createInstance();
    
            let opcServer = new opcda.OPCServer();
            await opcServer.init(comObject);

            let opcBrowser = await opcServer.getBrowser();
            let items = await opcBrowser.browseAllFlat();

            // don't need to await it, so we can return immediately
            opcBrowser.end()
                .then(() => opcServer.end())
                .then(() => comServer.closeStub())
                .catch(e => RED.log.error(`Error closing browse session: ${e}`));

            return items;
        }

        brosweItems().then(items => {
            res.json({ items });
        }).catch(err => {
            res.json(errorMessage(err));
            RED.log.error(errorMessage(err));
        });
    });

    /**
     * 
     * @param {object} config 
     */
    function OPCDAServer(config) {
        EventEmitter.call(this);
        const node = this;
        let isOnCleanUp = false;
        let reconnecting = false;

        RED.nodes.createNode(this, config);

        if (!this.credentials) {
            return node.error(RED._("opc-da.error.missingconfig"));
        }

        //init variables
        let status = 'unknown';
        let isVerbose = (config.verbose == 'on' || config.verbose == 'off') ? (config.verbose == 'on') : RED.settings.get('verbose');
        let connOpts = {
            address: config.address,
            domain: config.domain,
            username: this.credentials.username,
            password: this.credentials.password,
            clsid: config.clsid,
            timeout: config.timeout
        };
        let groups = new Map();
        let comSession, comServer, comObject, opcServer;

        function onComServerError(e) {
            node.error(errorMessage(e));
            node.warn("Trying to reconnect...");
            //setup().catch(onComServerError);
        }

        function updateStatus(newStatus) {
            if (status == newStatus) return;

            status = newStatus;
            groups.forEach(group => group.onServerStatus(status));
        }

        async function setup() {
            let comSession = new Session();
            comSession = comSession.createSession(connOpts.domain, connOpts.username, connOpts.password);
            comSession.setGlobalSocketTimeout(connOpts.timeout);

            comServer = new ComServer(new Clsid(connOpts.clsid), connOpts.address, comSession);
            //comServer.on('error', onComServerError);
            
            comServer.on('e_classnotreg', function(){
                node.error(RED._("opc-da.error.classnotreg"));
            });

            comServer.on("disconnected", function(){
                node.error(RED._("opc-da.error.disconnected"));
            })

            comServer.on("e_accessdenied", function() {
                node.error(RED._("opc-da.error.accessdenied"));
            });

            await comServer.init();
            
            comObject = await comServer.createInstance();

            opcServer = new opcda.OPCServer();
            await opcServer.init(comObject);
            
            for (const entry of groups.entries()) {
                const name = entry[0];
                const group = entry[1];
                let opcGroup = await opcServer.addGroup(name, group.opcConfig);
                console.log("setup for group: " + name);
                await group.updateInstance(opcGroup);
            }

            updateStatus('online');
        }

        async function cleanup() {
            console.log("asdfasdfasdfasdf");
            try {
                if (isOnCleanUp) return;
                console.log("Cleaning Up");
                isOnCleanUp = true;
                //cleanup groups first
                console.log("Cleaning groups...");
                for (const group of groups.values()) {
                    await group.cleanUp();
                }
                console.log("Cleaned Groups");
                if (opcServer) {
                    await opcServer.end();
                    opcServer = null;
                }
                console.log("Cleaned opcServer");
                if (comSession) {
                    await comSession.destroySession();
                    comServer = null;
                }
                console.log("Cleaned session. Finished.");
                isOnCleanUp = false;
            } catch (e) {
                //TODO I18N
                isOnCleanUp = false;
                let err = e && e.stack || e;
                console.log(e);
                node.error("Error cleaning up server: " + err, { error: err });
            }

            updateStatus('unknown');
        }

        node.reConnect = async function reConnect() {
            /* if reconnect was already called, do nothing
               if reconnect was never called, try to restart the session */
            if (!reconnecting) {
                reconnecting = true;
                await cleanup();
                await setup().catch(onComServerError);
                reconnecting = false
            }
        }

        node.getStatus = function getStatus() {
            return status;
        };

        node.registerGroup = function registerGroup(group) {
            if (groups.has(group.config.name)) {
                return RED._('opc-da.warn.dupgroupname');
            }

            groups.set(group.config.name, group);
        }

        node.unregisterGroup = function unregisterGroup(group) {
            groups.delete(group.config.name);
        }

        setup().catch(onComServerError);
    }
    RED.nodes.registerType("opc-da server", OPCDAServer, {
        credentials: {
            username: { type: "text" },
            password: { type: "password" }
        }
    });


    // ---------- OPC-DA Group ----------
    /**
     * @param {object} config 
     * @param {string} config.server
     * @param {string} config.updaterate
     * @param {string} config.deadband
     * @param {boolean} config.active
     * @param {boolean} config.validate
     * @param {object[]} config.vartable
     */
    function OPCDAGroup(config) {
        EventEmitter.call(this);
        const node = this;
        RED.nodes.createNode(this, config);

        node.server = RED.nodes.getNode(config.server);
        if (!node.server || !node.server.registerGroup) {
            return node.error(RED._("opc-da.error.missingconfig"));
        }

        /** @type {OPCGroupStateManager} */
        let opcGroupMgr;
        /** @type {OPCItemManager} */
        let opcItemMgr;
        /** @type {OPCSyncIO} */
        let opcSyncIo;
        let clientHandlePtr;
        let serverHandles = [], clientHandles = [];
        let status, timer;

        let readInProgress = false;
        let connected = false;
        let readDeferred = 0;
        let oldItems = {};
        let updateRate = parseInt(config.updaterate);
        let deadband = parseInt(config.deadband);
        let validate = config.validate;

        if (isNaN(updateRate)) {
            updateRate = 1000;
        }
        if (isNaN(deadband)) {
            deadband = 0;
        }

        node.config = config;
        node.opcConfig = {
            active: config.active,
            updateRate: updateRate,
            timeBias: 0,
            deadband: deadband || 0
        }

         /**
         * @private
         * @param {OPCGroupStateManager} newGroup
         */
        async function setup(newGroup) {
            clearInterval(timer);
            
            try {
                opcGroupMgr = newGroup;
                opcItemMgr = await opcGroupMgr.getItemManager();
                opcSyncIo = await opcGroupMgr.getSyncIO();
                
                clientHandlePtr = 1;
                clientHandles.length = 0;
                serverHandles = [];
                connected = true;
                readInProgress = false;
                readDeferred = 0;
                
                let items = config.vartable || [];
                if (items.length < 1) {
                    node.warn("opc-da.warn.noitems");
                }

                let itemsList = items.map(e => {
                    return { itemID: e.item, clientHandle: clientHandlePtr++ }
                });

                let resAddItems = await opcItemMgr.add(itemsList);
               
                for (let i = 0; i < resAddItems.length; i++) {
                    const resItem = resAddItems[i];
                    const item = itemsList[i];

                    if (resItem[0] !== 0) {
                        node.error(`Error adding item '${itemsList[i].itemID}': ${errorMessage(resItem[0])}`);
                    } else {
                        serverHandles.push(resItem[1].serverHandle);
                        clientHandles[item.clientHandle] = item.itemID;
                    }
                }
            } catch (e) {
                let err = e && e.stack || e;
                console.log(e);
                node.error("Error on setting up group: " + err);
            }

            // we set up the timer regardless the result of setting up items
            // we may support adding items at a later time
            if (updateRate < MIN_UPDATE_RATE) {
                updateRate = MIN_UPDATE_RATE;
                node.warn(RED._('opc-da.warn.minupdaterate', { value: updateRate + 'ms' }))
            }

            if (config.active) {
                timer = setInterval(doCycle, updateRate);
                doCycle();
            }
        }

        async function cleanup() {
            clearInterval(timer);
            clientHandlePtr = 1;
            clientHandles.length = 0;
            serverHandles = [];

            try {
                if (opcSyncIo) {
                    await opcSyncIo.end();
                    console.log("GroupCLeanup - opcSync");
                    opcSyncIo = null;
                }
                
                if (opcItemMgr) {
                    await opcItemMgr.end();
                    console.log("GroupCLeanup - opcItemMgr");
                    opcItemMgr = null;
                }
                
                if (opcGroupMgr) {
                    await opcGroupMgr.end();
                    console.log("GroupCLeanup - opcGroupMgr");
                    opcGroupMgr = null;
                }
            } catch (e) {
                let err = e && e.stack || e;
                console.log(e);
                node.error("Error on cleaning up group: " + err);
            }
        }

        async function doCycle() {
            if (connected && !readInProgress) {
                if (!serverHandles.length) return;

                readInProgress = true;
                readDeferred = 0;
                await opcSyncIo.read(opcda.constants.opc.dataSource.DEVICE, serverHandles)
                    .then(cycleCallback).catch(cycleError);
            } else {
                readDeferred++;
                if (readDeferred > 10) {
                    node.warn(RED._("opc-da.error.noresponse"), {});
                    clearInterval(timer);
                    // since we have no good way to know if there is a network problem
                    // or if something else happened, restart the whole thing
                    node.server.reConnect();
                }
            }
        }

        function cycleCallback(values) {
            readInProgress = false;
           
            if (readDeferred && connected) {
                doCycle();
                readDeferred = 0;
            }

            let changed = false;
            for (const item of values) {
                const itemID = clientHandles[item.clientHandle];
                
                if (!itemID) {
                    //TODO - what is the right to do here?
                    node.warn("Server replied with an unknown client handle");
                    continue;
                }

                let oldItem = oldItems[itemID];
                
                if (!oldItem || oldItem.quality !== item.quality || !equals(oldItem.value, item.value)) {
                    changed = true;
                    node.emit(itemID, item);
                    node.emit('__CHANGED__', { itemID, item });
                }
                oldItems[itemID] = item;
            }
            node.emit('__ALL__', oldItems);
            if (changed) node.emit('__ALL_CHANGED__', oldItems);
        }

        function cycleError(err) {
            readInProgress = false;
            node.error('Error reading items: ' + err && err.stack || err);
        }

        node.onServerStatus = function onServerStatus(s) {
            status = s;
            node.emit('__STATUS__', s);
        }

        node.getStatus = function getStatus() {
            return status;
        };

        node.cleanUp = async function cleanUp() {
            await cleanup();
        }

        /**
         * @private
         * @param {OPCGroupStateManager} newOpcGroup
         */
        node.updateInstance = async function updateInstance(newOpcGroup) {
            await cleanup();
            await setup(newOpcGroup);
        }

        node.on('close', async function (done) {
            node.server.unregisterGroup(this);
            await cleanup();
            console.log("group cleaned");
            done();
        });
        let err = node.server.registerGroup(this);
        if (err) {
            node.error(err, { error: err });
        }

    }
    RED.nodes.registerType("opc-da group", OPCDAGroup);


    // ---------- OPC-DA In ----------
    /**
     * @param {object} config
     * @param {string} config.group
     * @param {string} config.item
     * @param {string} config.mode
     * @param {boolean} config.diff
     */
    function OPCDAIn(config) {
        const node = this;
        RED.nodes.createNode(this, config);

        node.group = RED.nodes.getNode(config.group);
        if (!node.group || !node.group.getStatus) {
            return node.error(RED._("opc-da.error.missingconfig"));
        }

        let statusVal;

        function sendMsg(data, key, status) {
            // if there is no data to be sent
            if (!data) return;
            if (key === undefined) key = '';

            let msg;
            if (key === '') { //should be the case when mode == 'all'
                let newData = new Array();
                for (let key in data) {
                    newData.push({
                        errorCode: data[key].errorCode,
                        value: data[key].value,
                        quality: data[key].quality,
                        timestamp: data[key].timestamp,
                        topic: key
                    });
                }

                msg = {
                    topic: "all",
                    payload: newData
                };
            } else {
                if (data.errorCode !== 0) {
                    //TODO i18n and node status handling
                    msg = {
                        errorCode: data.errorCode,
                        payload: data.value,
                        quality: data.quality,
                        timestamp: data.timestamp,
                        topic: key
                    }
                    node.error(`Read of item '${key}' returned error: ${data.errorCode}`, msg);
                    return;
                }

                msg = {
                    payload: data.value,
                    quality: data.quality,
                    timestamp: data.timestamp,
                    topic: key
                };
            }
            statusVal = status !== undefined ? status : data;
            node.send(msg);
            node.status(generateStatus(node.group.getStatus(), statusVal));
        }

        function onChanged(elm) {
            sendMsg(elm.item, elm.itemID, null);
        }

        function onDataSplit(data) {
            Object.keys(data).forEach(function (key) {
                sendMsg(data[key], key, null);
            });
        }

        function onData(data) {
            sendMsg(data, config.mode == 'single' ? config.item : '');
        }

        function onDataSelect(data) {
            onData(data[config.item]);
        }

        function onGroupStatus(s) {
            node.status(generateStatus(s.status, statusVal));
        }

        node.group.on('__STATUS__', onGroupStatus);
        node.status(generateStatus(node.group.getStatus(), statusVal));

        if (config.diff) {
            switch (config.mode) {
                case 'all-split':
                    node.group.on('__CHANGED__', onChanged);
                    break;
                case 'single':
                    node.group.on(config.item, onData);
                    break;
                case 'all':
                default:
                    node.group.on('__ALL_CHANGED__', onData);
            }
        } else {
            switch (config.mode) {
                case 'all-split':
                    node.group.on('__ALL__', onDataSplit);
                    break;
                case 'single':
                    node.group.on('__ALL__', onDataSelect);
                    break;
                case 'all':
                default:
                    node.group.on('__ALL__', onData);
            }
        }

        node.on('close', function (done) {
            node.group.removeListener('__ALL__', onDataSelect);
            node.group.removeListener('__ALL__', onDataSplit);
            node.group.removeListener('__ALL__', onData);
            node.group.removeListener('__ALL_CHANGED__', onData);
            node.group.removeListener('__CHANGED__', onChanged);
            node.group.removeListener('__STATUS__', onGroupStatus);
            node.group.removeListener(config.item, onData);
            done();
        });
    }
    RED.nodes.registerType("opc-da in", OPCDAIn);


    // ---------- OPC-DA Out ----------
    /**
     * 
     * @param {object} config
     * @param {string} config.group
     * @param {string} config.item
     */
    function OPCDAOut(config) {
        const node = this;
        RED.nodes.createNode(this, config);

        node.group = RED.nodes.getNode(config.group);
        if (!node.group) {
            return node.error(RED._("opc-da.error.missingconfig"));
        }

        let statusVal;

        function onGroupStatus(s) {
            node.status(generateStatus(s.status, statusVal));
        }

        function onNewMsg(msg) {
            var writeObj = {
                name: config.item || msg.item,
                val: msg.payload
            };

            if (!writeObj.name) return;

            statusVal = writeObj.val;
            node.group.writeVar(writeObj);
            node.status(generateStatus(node.group.getStatus(), statusVal));
        }

        node.status(generateStatus(node.group.getStatus(), statusVal));

        node.on('input', onNewMsg);
        node.group.on('__STATUS__', onGroupStatus);

        node.on('close', function (done) {
            node.group.removeListener('__STATUS__', onGroupStatus);
            done();
        });

    }

    /**
     * @private
     * @param {Number} errorCode 
     */
    function errorMessage(errorCode) {
        let msgText;
        
        switch(errorCode){
            case 0x00000005:
                msgText = "Access denied. Username and/or password might be wrong."
                break;
            case 0xC0040006:
                msgText = "The Items AccessRights do not allow the operation.";
                break;
            case 0xC0040004:
                msgText =  "The server cannot convert the data between the specified format/ requested data type and the canonical data type.";
                break;
            case 0xC004000C:
                msgText = "Duplicate name not allowed.";
                break;
            case 0xC0040010: 	
                msgText = "The server's configuration file is an invalid format.";
                break;
            case 0xC0040009: 
                msgText = "The filter string was not valid";
                break;
            case 0xC0040001: 
                msgText = "The value of the handle is invalid. Note: a client should never pass an invalid handle to a server. If this error occurs, it is due to a programming error in the client or possibly in the server.";
                break;
            case 0xC0040008: 
                msgText = "The item ID doesn't conform to the server's syntax.";
                break;
            case 0xC0040203: 
                msgText = "The passed property ID is not valid for the item.";
                break;
            case 0xC0040011: 
                msgText = "Requested Object (e.g. a public group) was not found.";
                break;
            case 0xC0040005: 
                msgText = "The requested operation cannot be done on a public group.";
                break;
            case 0xC004000B: 
                msgText = "The value was out of range.";
                break;
            case 0xC0040007: 
                msgText = "The item ID is not defined in the server address space (on add or validate) or no longer exists in the server address space (for read or write).";
                break;
            case 0xC004000A: 
                msgText = "The item's access path is not known to the server.";
                break;
            case 0x0004000E: 
                msgText = "A value passed to WRITE was accepted but the output was clamped.";
                break;
            case 0x0004000F: 
                msgText = "The operation cannot be performed because the object is being referenced.";
                break;
            case 0x0004000D: 
                msgText = "The server does not support the requested data rate but will use the closest available rate.";
                break;
        }
        return String(errorCode) + " - " + msgText;
    }
    RED.nodes.registerType("opc-da out", OPCDAOut);
};
