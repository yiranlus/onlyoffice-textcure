# TextCure

An unofficial Druide Antidote plugin for OnlyOffice.

# Features

* Support all major types of document of OnlyOffice
* Live correction with Antidote

## Installation

You will of course need a license of [Druide Antidote](https://www.antidote.info/) to use this plugin. This plugin depends on the local installation of [Connectix](https://www.antidote.info/fr/integrations/documentation/connectix/utilitaire) that comes with Antidote.

You can download the latest version from the [release page](https://github.com/yiranlus/textcure/releases) of the project on GitHub.

You should be able to find the file ending with `.plugin` in each released version. Download it and install it manually to OnlyOffice.

Currently, the plugin works well with Desktop Editor. I did not test it with the OnlyOffice in the browser. I may publish this plugin to the OnlyOffice Plugin Marketplace some day.

## Build

If you are interested in building the plugin yourself, you can fork the project on GitHub and clone it to your PC. The project is written in Typescript to avoid some Javascript quirks.

First, you need to install the dependencies using:

```
npm install
```

in the project root folder.

Then, you can use the following command to build the plugin:

```
npm run build:plugin
```

After that, you will be able to find a folder named `textcure` that contains all the compiled Javascript files and another file name ending with `.plugin` at the project root.
