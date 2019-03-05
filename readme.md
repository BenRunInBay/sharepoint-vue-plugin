# SharePoint Vue Plug-in
## Last updated 2019-03-05

[GitHub repo](https://github.com/BenRunInBay)

## Vue main.js entry:
```javascript
import SharePoint from '@/lib/SharePoint'
Vue.use(SharePoint, '/sites/MySite/')
````

Or use on a page:
```html
<script src="axios.min.js"></script>
<script type="module">
    import SharePoint from "./lib/SharePoint.VuePlugin.js";
    Vue.use(SharePoint, '/sites/MySite/')
</script>
```

Within the Vue app or component, refer to the SharePoint as this.$sp
```javascript
created() {
    this.$sp.getRequestDigest(() => {
    // Once completed, you can start recording data to a list by specifying list name
    // and an object containing the column names of that list that you want to fill.
    this.$sp.addListItem({
        listName: 'MyList',
        itemData: { Title: 'title', CustomColumn: 'custom value' },
    });
    })
}
````