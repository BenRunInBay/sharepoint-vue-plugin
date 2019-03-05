# SharePoint Vue Plug-in
## Last updated 2019-03-05

## Vue main.js entry:
```javascript
import SharePoint from '@/lib/SharePoint'
Vue.use(SharePoint, '/sites/MySite/')
````

Within the Vue app or component, refer to the SharePoint as this.$sp
```javascript
created() {
    // read data from a list immediately
    this.$sp.getList({
        listName: 'MyList',
        select: "ID,Title
        });
    // to write data, you have obtain a digest value first
    this.$sp.getFormDigest(() => {
        // Once the digest is known, you can start writing or updating list data
        this.$sp.addListItem({
            listName: 'MyList',
            itemData: { Title: 'title', CustomColumn: 'custom value' },
            });
    })
}
````

[GitHub repo](https://github.com/BenRunInBay)