# SharePoint Vue Plug-in
## Last updated 2019-03-05

## Vue main.js entry:
```javascript
import 'sharepoint-vue-plugin'
// specify path to the SharePoint site are using this within
Vue.use(SharePoint, '/sites/MySite/')
````

Within the Vue app or component, refer to the SharePoint as this.$sp
```javascript
data: () => { return {
    myVueList: []
}},
created() {
    // read data from a list immediately
    this.$sp.getList({
        listName: 'MyList',
        select: "ID,Title
        })
        .then((listData) => {
            listData.forEach(item => {
                this.myVueList.push(item);
            })
        });
    // to write data, you have obtain a digest value first
    this.$sp.getFormDigest(() => {
        // Once the digest is known, you can start writing or updating list data
        this.$sp.addListItem({
            listName: 'MyList',
            itemData: {
                Title: 'Updated title',
                CustomColumn: 'custom value' }
            });
    })
}
````

[GitHub repo](https://github.com/BenRunInBay/sharepoint-vue-plugin/)