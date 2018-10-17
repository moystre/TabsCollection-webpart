# WorkPointUI

This repository contains the WorkPoint user interface built using the SharePoint Framework.

## Seeing it in action

In the development environment, the app and source files are default hosted at "https://wp365devmodernui.azureedge.net".

This ensures that backend developers and other visitors to development WorkPoint365 solutions can see the WorkPoint UI functionality.

## Debugging and developing

If custom development or debugging is required, the developer needs to build, bundle and package the code, as described below. Be careful not to include the --ship or --prod flags to any of these commands during development.

When this is done, the resulting "workpointui.sppkg" app file (located here: /sharepoint/solution/workpointui.sppkg) needs to be uploaded to the site collection app catalog.

Overwrite the already existing app, and confirm the "Install to entire solution..." text.

Afterwards check-in the app file, and the app will now point to the locally hosted files (https://localhost:4321/<...>), served via. the "gulp serve --nobrowser" command.

## Building the code

In your preferred terminal, console or command promt, run the following commands in the root of this project:

1. Clone the repo using git (https://workpoint365.visualstudio.com/WorkPoint%20Online/_git/WorkPointUI)
````bash
git clone https://workpoint365.visualstudio.com/WorkPoint%20Online/_git/WorkPointUI
````

2. Install gulp, if not already installed ( -g is optional)
````bash
npm install -g gulp
````

````
3. Install dependencies
````bash
npm install

4. Build the project
````bash
gulp build
````

Bulding this package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

#### Note that these directories are NOT checked in to the version control!

## Running the code
1. Trust the development certificate
````bash
gulp trust-dev-cert
````

2. Serve the project without opening a browser
````bash
gulp serve --nobrowser
````


Errors? Try opening your browser console window.
You might have forgotten to trust the development certificate.

## Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO

# Localizations
Add follwing property and value in the 'write-manifests.json' file:

````JavaScript
"debugLocale": "da-dk" // Or other locale
````

# WorkPointUI Build configurations

## CDN setup
config/write-maifests.json: cdnBasePath: "https://wp365modernui.azureedge.net/WorkPointUI/"


## Product id
* Development:                              ef77d173-14f4-4806-a005-3a799acc4698
* Test:                                     8767526e-5996-4bd1-ae34-756e044d6c42
* Production:                               a75dc1df-acaa-4dee-a7a4-f48baa076999

## Aliases and ClientSideComponentIds

### Development
* WorkPointApplicationCustomizer:           02a45ec8-0832-4ccb-8378-82c2ddfb3542
* SiteFieldCustomizer:                      917a98fc-71fd-4761-ad49-00ec4ae12b28
* TemplateLibraryCommandSet:                f39076f4-51d8-4fec-9d4c-885f25246efb
* TemplateLibraryFieldCustomizer:           ba7bda5a-e3c5-434c-a44c-f7e9193c7862
* EmailManagerWebPart:                      7cb5142e-2fbc-4bbe-8259-d16e7156125e
* JournalWebPart:                           c0deaf4c-d346-4fde-8c2c-bbccd2d7e39b
* RiskMatrixWebPart:                        96460fd6-e55b-4ad4-a4bb-f18c71f2c75a
* TaskOverviewWebPart:                      8725b854-5f2e-4da8-a3ee-08eea4d68852
* RelationWebpart:                          6c8c95d9-ec08-4a97-9330-ebfd7773db11
* WorkPointListViewWebPart:                 7894de77-b834-4d03-a536-cb3ee094b1f8
* SiteManagementCommandSet:                 0ddf279d-6a54-46ef-aab5-6bdca1a58ae3
* MasterSiteSyncLogFieldCustomizer:         039f1e46-eeca-4f82-83bd-c432ed5228a9
* MasterSiteSyncResultFieldCustomizer:      9b1fda6d-8101-418b-b489-c1e70afbcfbb
* SecurityReplicationResultFieldCustomizer: 3903babb-bc1e-492c-bfb9-c56a93607b21
* SecurityReplicationLogFieldCustomizer:    128c303b-5a79-4dc0-a9bc-57fa588e6136

### Test
* WorkPointApplicationCustomizer:           fbdf968d-66c8-414c-9329-fafb68c95f4d
* SiteFieldCustomizer:                      02ac6bf6-1980-4b86-a0cb-d9ef44868cbe
* TemplateLibraryCommandSet:                f8420cd9-9656-432a-9a92-12fe16023dce
* TemplateLibraryFieldCustomizer:           d5f56132-a0f3-4403-9d0d-86153573491a
* EmailManagerWebPart:                      be72ce65-4e3d-4845-8ac6-0b244ce480d9
* JournalWebPart:                           597c14f1-41a5-400c-87e2-b6168db87c85
* RiskMatrixWebPart:                        fd9d259f-46f0-43bd-b26e-7a9af4f00a8b
* TaskOverviewWebPart:                      c2a7ce9e-b79c-491a-89b3-88bda6b615fe
* RelationWebpart:                          f0f92abe-6cee-4428-9b61-b1ecd2aaa81c
* WorkPointListViewWebPart:                 a9722ee7-7b58-4c9d-8746-16f31317fc6d
* SiteManagementCommandSet:                 368e8ff1-c83a-4ccd-9376-a11cf67e37e8
* MasterSiteSyncLogFieldCustomizer:         85e5e41a-8de8-46b0-b25f-1e98447673f8
* MasterSiteSyncResultFieldCustomizer:      78377cc3-7cd1-48fe-8d5c-3a59febdb9cc
* SecurityReplicationResultFieldCustomizer: 732fcd28-c07f-43d8-b06e-b1fefa9f144e
* SecurityReplicationLogFieldCustomizer:    2b46c3c2-8a27-42cb-a585-be0dee921816

### Production
* WorkPointApplicationCustomizer:           12ce863a-6701-4469-a180-82b3a06ee05f
* SiteFieldCustomizer:                      fb4df2df-5b35-4219-b0a5-3e0ba9b25405
* TemplateLibraryCommandSet:                a190b131-6104-45e3-b0f9-03abc14142c1
* TemplateLibraryFieldCustomizer:           47a4d0cd-8b51-4bd8-bff7-aa3b7cb1b870
* EmailManagerWebPart:                      514a2006-a412-4116-bd35-ca0adeaec040
* JournalWebPart:                           b30f88d2-3766-4d58-ba46-bf46563ff51e
* RiskMatrixWebPart:                        4af944db-b86a-494a-ba25-531c8287db2d
* TaskOverviewWebPart:                      f5502e94-76a3-44ae-be35-76f2aab95f07
* RelationWebpart:                          21572e50-e64d-4dbc-89ce-bbd32bde2d21
* WorkPointListViewWebPart:                 14da319d-c91c-4410-9d6d-b799b824915b
* SiteManagementCommandSet:                 51382e83-10ec-41b3-9d26-a2f6f58b3aca
* MasterSiteSyncLogFieldCustomizer:         08059d7d-61b5-46f4-abe7-4cc889c3aafb
* MasterSiteSyncResultFieldCustomizer:      8a17deee-5c56-47b0-aa16-5629d10df6ce
* SecurityReplicationResultFieldCustomizer: 51125935-b245-4d3d-9688-5ed2afe35134
* SecurityReplicationLogFieldCustomizer:    796bf42c-e2f1-4a3f-b037-c1fe5da7dbda


# Web APIs
By exchanging the content in "/config/webapi-config.json" with below values, a different web API can be requested.

## Dev (Default)
````JavaScript
[
    {
        "name": "WorkPoint365",
        "id": "8129904a-6925-4414-8d10-6d1d338a1b84",
        "url": "https://localhost:44302"
    },
    {
        "name": "EMM365",
        "id": "2d1103ee-a86e-4501-83ae-5f1822cf368f",
        "url": "https://emm365webapidev.azurewebsites.net/api/v1/"
    }
]
````

## Test
````JavaScript
[
    {
        "name": "WorkPoint365 WebAPI Test",
        "id": "d95bdcb0-b057-41c8-88e2-efd4936e8b32",
        "url": "https://wp365webapitest.azurewebsites.net"
    },
    {
        "name": "EMM365.WebAPI.Test",
        "id": "bb157834-4656-4cd7-b32f-3ea75e84aa55",
        "url": "https://emm365webapi-test.azurewebsites.net/"
    }
]
````

# Troubleshooting

## Slow build/bundle/serve in development?
SpFx is notoriously slow at running multiple web parts/extensions at a time. The development command `gulp serve --nobrowser` takes long when running the 'webpack' subtask (in some instances this has been seen between 35 - 50 seconds - 19/07-2018) on a full project.

To alleviate this, the developer can comment out sections of the `config.json` file located in the config folder. 

> Warning: We always recommend not touching the 'work-point-application-customizer' part, as this is the primary user interface for WorkPoint.

Comment out web parts and extensions that are not vital for the current development purposes, and you will see a dramatic drop in webpack build time.

> Hint: If you perform a full build/bundle/serve command, before commenting out the config parts, the already built components will still function. This will of course be negated if the developer runs the `gulp clean` command.

> **VITAL: Do not check-in the `config.json` file with commented out sections! This will ruin the production build pipeline by only building the web parts / extensions that are not commented out.**#   W P 3 6 5 - t a b u l a t o r - w e b p a r t  
 #   T a b s C o l l e c t i o n - w e b p a r t  
 