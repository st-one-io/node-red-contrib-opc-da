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
            //console.log('browseItems', params.address, params.domain, params.username, params.password, params.clsid)
            let { comServer, opcServer } = await opcda.createServer(params.address, params.domain, params.username, params.password, params.clsid);

            let opcBrowser = await opcServer.getBrowser();
            let items = await opcBrowser.browseAllFlat();

            // don't need to await it, so we can return immediately
            opcBrowser.end()
                .then(() => opcServer.end())
                .then(() => comServer.closeStub())
                .catch(e => RED.log.error(`Error closing browse session: ${e}`));

            return items;
        }

        brosweItems().then(items => res.json({ items })).catch(err => {
            res.json({ err: err.toString() });
            RED.log.info(err);
        });
    });

    /**
     * 
     * @param {object} config 
     */
    function OPCDAServer(config) {
        EventEmitter.call(this);
        const node = this;

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
            clsid: config.clsid
        };
        let groups = new Map();
        let comSession, comServer, comObject, opcServer;

        function onComServerError(e) {
            //TODO improve this
            console.log(e);
            node.error(e && e.stack || e, {});
        }

        function updateStatus(newStatus) {
            if (status == newStatus) return;

            status = newStatus;
            groups.forEach(group => group.onServerStatus(status));
        }

        async function setup() {
            let comSession = new Session();
            comSession = comSession.createSession(connOpts.domain, connOpts.username, connOpts.password);

            comServer = new ComServer(new Clsid(connOpts.clsid), connOpts.address, comSession);
            //comServer.on('error', onComServerError);
            await comServer.init();

            comObject = await comServer.createInstance();

            opcServer = new opcda.OPCServer();
            await opcServer.init(comObject);

            for (const entry of groups.entries()) {
                const name = entry[0];
                const group = entry[1];
                console.log("ENTRY", entry);
                let opcGroup = await opcServer.addGroup(name, group.opcConfig);
                group.updateInstance(opcGroup);
            }

            updateStatus('online');
        }

        async function cleanup() {
            try {
                //cleanup groups first
                for (const group of groups.values()) {
                    await group.cleanup();
                }

                if (opcServer) {
                    await opcServer.end();
                    opcServer = null;
                }
                if (comServer) {
                    await comServer.closeStub();
                    comServer = null;
                }
            } catch (e) {
                //TODO I18N
                let err = e && e.stack || e;
                console.log(e);
                node.error("Error cleaning up server: " + err, { error: err });
            }

            updateStatus('unknown');
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
                        //TODO - get cause and I18N
                        node.error(`Error adding item '${itemsList[i].itemID}': ${resItem[0]}`);
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

            timer = setInterval(doCycle, updateRate);
            doCycle();
        }

        async function cleanup() {
            clearInterval(timer);
            clientHandlePtr = 1;
            clientHandles.length = 0;
            serverHandles = [];

            try {
                if (opcSyncIo) {
                    await opcSyncIo.end();
                    opcSyncIo = null;
                }
                if (opcItemMgr) {
                    await opcItemMgr.end();
                    opcItemMgr = null;
                }
                if (opcGroupMgr) {
                    await opcGroupMgr.end();
                    opcGroupMgr = null;
                }
            } catch (e) {
                let err = e && e.stack || e;
                console.log(e);
                node.error("Error on cleaning up group: " + err);
            }
        }

        function doCycle() {
            if (connected && !readInProgress) {
                if (!serverHandles.length) return;

                readInProgress = true;
                readDeferred = 0;
                opcSyncIo.read(opcda.constants.opc.dataSource.DEVICE, serverHandles)
                    .then(cycleCallback).catch(cycleError);
            } else {
                readDeferred++;
                if (readDeferred > 10) {
                    node.warn(RED._("opc-da.error.noresponse"), {});
                    //TODO - reset communication?
                }
            }
        }

        function cycleCallback(values) {
            readInProgress = false;

            if (readDeferred && connected) {
                doCycle();
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

            //TODO error handling
            console.log(err);
            node.error('Error reading items: ' + err && err.stack || err);
        }

        node.onServerStatus = function onServerStatus(s) {
            status = s;
            node.emit('__STATUS__', s);
        }

        node.getStatus = function getStatus() {
            return status;
        };

        /**
         * @private
         * @param {OPCGroupStateManager} newOpcGroup
         */
        node.updateInstance = function updateInstance(newOpcGroup) {
            cleanup().then(() => setup(newOpcGroup));
        }

        node.on('close', async function (done) {
            node.server.unregisterGroup(this);
            await cleanup();
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
            if (key === undefined) key = '';

            let msg;
            if (key === '') { //should be the case when mode == 'all'
                msg = data;
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
    RED.nodes.registerType("opc-da out", OPCDAOut);
};
