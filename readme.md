# SharePoint Vue Plug-in
### Last updated 2019-03-26
### NOT RELEASED, UNDER DEVELOPMENT

## Presentation about using Vue in SharePoint
[Vue in SharePoint](https://1drv.ms/p/s!AjelfXJUND_KgrATrHzrTwwRoG2X2A)

## Purpose
This is a compact Vue plug-in for accessing SharePoint REST API to perform list/library CRUD operations, identifying current user, getting profile data of other users in the system, and sending emails through the SharePoint site. It does NOT require any other SharePoint client libraries. You can install this in your vue-cli 3 webpack-built application and use it independently.

## Requirements
- Vue
- ES2015 or build-environment with polyfills
- axios

## Methods
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
Most methods return a Promise.

## Vue main.js entry:
```javascript
import 'sharepoint-vue-plugin'
// specify path to the SharePoint site that you are using this within
Vue.use(SharePoint, '/sites/MySite/')
````
## Using it
Within the Vue app or component, refer to the SharePoint object as this.$sp
```javascript
data: () => { return {
    myVueList: [],
    errorMessage: ""
}},
created() {
    // read data from a list immediately
    this.$sp.getList({
        listName: 'MyList',
        select: "ID,Title"
        }).then(listData => {
            listData.forEach(item => {
                this.myVueList.push(item);
            })
        }).catch(error => {
            this.errorMessage = error;
        });
    // to write data, you have to obtain a digest value first
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
````

[GitHub repo](https://github.com/BenRunInBay/sharepoint-vue-plugin/)