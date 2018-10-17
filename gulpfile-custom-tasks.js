'use strict';

const build = require('@microsoft/sp-build-web');
const fs = require('fs');
const log = require('fancy-log');
const stripJsonComments = require('strip-json-comments');

build.task('update-manifest', {
    execute: (config) => {
        return new Promise((resolve, reject) => {
            const cdnPath = config.args['cdnpath'] || "";
            if (cdnPath !== "") {
                let json = JSON.parse(fs.readFileSync('./config/write-manifests.json'));
                json.cdnBasePath = cdnPath;
                fs.writeFileSync('./config/write-manifests.json', JSON.stringify(json));
            }
            resolve();
        });
    }
});


build.task('update-package-solution', {
    execute: (config) => {
        return new Promise((resolve, reject) => {
            const wp365WebAPIResource = config.args['wp365webapiresource'] || "";
            const emm365WebAPIResource = config.args['emm365webapiresource'] || "";
            const solutionName = config.args['solutionname'] || "";
            const appProductId = config.args['appproductid'] || "";
            if (wp365WebAPIResource !== "" && emm365WebAPIResource !== "") {
                let json = JSON.parse(fs.readFileSync('./config/package-solution.json'));
                json.solution.name = solutionName;
                json.solution.id = appProductId; 
                json.solution.webApiPermissionRequests =  [
                    { resource: wp365WebAPIResource, scope: "user_impersonation" },
                    { resource: emm365WebAPIResource, scope: "user_impersonation" },
                    { resource: "Windows Azure Active Directory", scope: "User.Read" }
                ];
                fs.writeFileSync('./config/package-solution.json', JSON.stringify(json));
            }
            resolve();
        });
    }
});

build.task('update-azure-storage-config', {
    execute: (config) => {
        return new Promise((resolve, reject) => {
            const account = config.args['account'] || "";
            const accessKey = config.args['accesskey'] || "";
            if (account !== "" && accessKey != "") {
                let json = JSON.parse(fs.readFileSync('./config/deploy-azure-storage.json'));
                json.account = account;                
                json.accessKey = accessKey;
                fs.writeFileSync('./config/deploy-azure-storage.json', JSON.stringify(json));
            }
            resolve();
        })
    }
})

build.task('update-webapi-config', {
    execute: (config) => {
        return new Promise((resolve, reject) => {
            const wp365WebAPIId = config.args['wp365webapiid'] || "";
            const wp365WebAPIUrl = config.args['wp365webapiurl'] || "";
            const emm365WebAPIId = config.args['emm365webapiid'] || "";
            const emm365WebAPIUrl = config.args['emm365webapiurl'] || "";

            if (wp365WebAPIId !== "" && wp365WebAPIUrl !== "" && emm365WebAPIId != "" && emm365WebAPIUrl != "") {
                let json = JSON.parse(fs.readFileSync('./config/webapi-config.json'));
                for (let i = 0; i < json.length; i++) {
                    let api = json[i];
                    if (api.name === "WorkPoint365") {
                        api.id = wp365WebAPIId;
                        api.url = wp365WebAPIUrl;
                    }
                    else if (api.name == "EMM365") {
                        api.id = emm365WebAPIId;
                        api.url = emm365WebAPIUrl;
                    }
                }
                fs.writeFileSync('./config/webapi-config.json', JSON.stringify(json));
            }
            resolve(); 
        })
    }
})


build.task('update-clientsidecomponentids', {
    execute: (config) => {
        return new Promise((resolve, reject) => {
            let appconfig = JSON.parse(fs.readFileSync('./config/config.json'));                
            for (let key in appconfig.bundles) {
                let bundle = appconfig.bundles[key];
                for (let i = 0; i < bundle.components.length; i++) {
                    let manifestsrc = bundle.components[i].manifest;
                    let manifest = JSON.parse(stripJsonComments(fs.readFileSync(manifestsrc, 'utf8')));
                    let clientsidecomponentid = config.args[manifest.alias] || "";
                    if (clientsidecomponentid !== "") {
                        manifest.id = clientsidecomponentid;
                        fs.writeFileSync(manifestsrc, JSON.stringify(manifest));
                    }
                }
            }
            resolve(); 
        })
    }
})

