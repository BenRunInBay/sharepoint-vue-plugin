# SharePoint Vue Plug-in

### Last updated 2019-06-02

## Purpose

This is a compact Vue plug-in for accessing the SharePoint REST API to perform list/library CRUD operations, identifying current user, getting profile data of other users in the system, and sending emails through the SharePoint site. It does NOT require any other SharePoint client libraries. You can install this in your vue-cli 3 webpack-built application and use it independently.

This custom Vue plug minifies to about 16k and is accessible within Vue components as:

```
this.$sp
```

### Alternatives

You could use the [SharePoint client library](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-javascript-library-code-in-sharepoint) or the [PnPjs library](https://pnp.github.io/pnpjs/). The SP client library is included when developing pages within a standard SharePoint master. The PnPjs library is about 40k.

## Requirements

- Vue
- ES2015 or build-environment with polyfills
- axios

## Installation

I have not created this into an NPM package yet. For now, just copy it to:
src/plugins/sharepoint-vue-plugin

## Vue main.js entry:

```javascript
import SharePoint from "./plugins/sharepoint-vue-plugin";
// specify path to the SharePoint site that you are using this within
Vue.use(SharePoint, "/sites/MySite/", {
  // specify the production domain of your SharePoint environment
  productionHosts: ["myhost.sharepoint.com"]
});
```

## Using it

Within the Vue app or component, refer to the SharePoint object as this.\$sp

```javascript
data: () => { return {
    myVueList: [],
    errorMessage: ""
}},
created() {
    // read data from a list immediately
    this.$sp.getListData({
        listName: 'MyList',
        select: "ID,Title"
        }).then(listData => {
            listData.forEach(item => {
                this.myVueList.push(item);
            })
        }).catch(error => {
            this.errorMessage = error;
        });
    // to write data, you have to obtain a form digest value first
    this.$sp.getFormDigest().then(() => {
        // Once the digest is known, you can start writing or updating list data
        this.$sp.addListItem({
            listName: 'MyList',
            itemData: {
                Title: 'Updated title',
                CustomColumn: 'custom value' }
            }).catch(error => {
                this.errorMessage = error;
            });
    })
}
```

## Methods

Most methods return a Promise.

- getFormDigest()
- isWriteReady()
- getListData({ listName, select, filter, expand, orderby, top, devStaticDataUrl })
- addListItem({ listName, itemData })
- updateListItem({ listName, itemData, itemUrl })
- deleteListItem({ itemUrl })
- retrievePeopleProfile({ accountName, email = null, property = null })
- retrieveCurrentUserProfile()
- ensureSiteUserId({ accountName })
- sendEmail({ from, to, subject, body })
- get({ baseUrl, path, url, devStaticDataUrl })
- post({ path, url, data })

[GitHub repo](https://github.com/BenRunInBay/sharepoint-vue-plugin/)
