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

module.exports = function (RED) {

    const EventEmitter = require('events').EventEmitter;
    const opcda = require('opc-da');

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
            let {comServer, opcServer} = await opcda.createServer(params.address, params.domain, params.username, params.password, params.clsid);

            let opcBrowser = await opcServer.getBrowser();
            let items = await opcBrowser.browseAllFlat();

            // TODO - close and cleanup everything

            return items;
        }

        brosweItems().then(items => res.json({items})).catch(err => {
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

        node.getStatus = function getStatus() {
            return status;
        };

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
     * @param {number} config.updaterate
     * @param {number} config.deadband
     * @param {boolean} config.active
     * @param {boolean} config.validate
     * @param {object[]} config.vartable
     */
    function OPCDAGroup(config) {
        EventEmitter.call(this);
        const node = this;
        RED.nodes.createNode(this, config);

        node.server = RED.nodes.getNode(config.server);
        if (!node.server) {
            return node.error(RED._("opc-da.error.missingconfig"));
        }

        let updateRate = config.updaterate || 1000;
        let timeBias = 0; //hardcoded for now
        let deadband = config.deadband || 0;
        let active = config.active;
        let vartable = config.vartable;

        let status;

        node.getStatus = function getStatus() {
            return status;
        };

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
        if (!node.group) {
            return node.error(RED._("opc-da.error.missingconfig"));
        }

        let statusVal;

        function sendMsg(data, key, status) {
            if (key === undefined) key = '';
            var msg = {
                payload: data,
                topic: key
            };
            statusVal = status !== undefined ? status : data;
            node.send(msg);
            node.status(generateStatus(node.group.getStatus(), statusVal));
        }

        function onChanged(variable) {
            sendMsg(variable.value, variable.key, null);
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
