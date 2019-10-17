# node-red-contrib-opc-da

node-red-contrib-opc-da is an OPC-DA compatible node for Node-RED that allow interaction with remote OPC-DA servers. Currently only reading and browsing operations are supported.

## Table of Contents

- [Install](#install)
- [Usage]()
  - [Creating a Server](#creating-a-server)
  - [Creating a Group](#creating-a-group)
  - [Adding Items to a Group](#adding-items-to-a-group)
- [Contributing](#contributing)

## Install

Using npm:

```bash
npm install node-red-contrib-opc-da
```

## Creating a Server

To create a server you will need a few information about your target server: the IP address, the domain name, a username with enough privilege to remotely interact with the OPC Server, this users's password, and a [CLSID](https://docs.microsoft.com/en-us/windows/win32/com/clsid). We ship this node with a few known [ProgIds](https://docs.microsoft.com/en-us/windows/win32/com/-progid--key), which will fill the CLSID field with the correct string. If you have one or more applications that you think could be included on the default options feel free to open an issue with your suggestion. In case your server ProgId is not listed, you can choose the ```Custom``` options and type it by hand.

![](/images/createserver.png)

You should also pay attention to the timeout value to make sure it is compatible with the characteristics of your network. If this value is too low, the server might not even be created, and if does other problems related to timeouts might arise. Finally, if you want to test your configuration click the ```Test and get Items``` button. This button will connect to the server, authenticate, and will browse for a full list of available items.

## Creating a Group

Once your server was created you'll have to create a group. For a group to be created you must first select a server you previously created. The update rate defines how frequent the server will be queried for the items added to this group. You can also use the ```Active``` option to activate or deactivate your groups. For now, the ```Deadband``` feature is not fully implemented so you don't need to bother with it.

![](/images/creategroups.png)

## Adding Items to a Group

To add Items to a group, you can type the item name as it is stored at the server and click the ```+``` button. In case you are not sure which items are available on your server, return to the OPC Server configuration tab and click the ```Test and get Items``` button since it will browse and return a list of full available items, allowing you to add from a list here.

| ![](/images/additem01.png) | ![](/images/additem02.png) |
| :------------------------: | -------------------------- |
|                            |                            |
## Contributing

This is a partial implementation and there are lots that could be done to improve what is already supported or to add support for more OPC-DA features. Feel free to dive in! Open an issue or submit PRs.
