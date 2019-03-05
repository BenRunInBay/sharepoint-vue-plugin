/*
    SharePoint Vue Plug-in
    https://github.com/BenRunInBay

    Last updated 2019-03-05

    Vue main.js entry:
        import SharePoint from '@/lib/SharePoint'
        Vue.use(SharePoint, '/sites/MySite/')

    Or use on a page:
        //comment out axios module import below
        <script src="axios.min.js"></script>
        <script type="module">
          import SharePoint from "./lib/SharePoint.VuePlugin.js";
          Vue.use(SharePoint, '/sites/MySite/')
        </script>

    Within the Vue app or component, refer to the SharePoint as this.$sp
        created() {
          this.$sp.getFormDigest(() => {
            // Once completed, you can start recording data to a list by specifying list name
            // and an object containing the column names of that list that you want to fill.
            this.$sp.addListItem({
              listName: 'MyList',
              itemData: { Title: 'title', CustomColumn: 'custom value' },
            });
          })
        }
*/
/* Comment out the following line if not using this in a webpack build system: */
import axios from "axios";

// artificial delay when working in Dev mode
const constDevLoadDelay = 1000;

/* Modify these default configurations for your SharePoint environment */
let baseConfig = {
    productionHosts: ["yoursite.sharepoint.com"],
    showConsoleActivityInDev: true,
    profileDefaultSelect:
      "AccountName,DisplayName,Email,PictureUrl,PersonalUrl,Title,UserProfileProperties",
    myProfileDefaultSelect:
      "DisplayName,AccountName,Email,PictureUrl,PersonalUrl,Title",
    listPath: "_api/Web/Lists/",
    currentUserPropertiesPrefix:
      "/_api/sp.userprofiles.peoplemanager/getmyproperties/?$select=",
    peopleManagerbaseUrl:
      "/_api/sp.userprofiles.peoplemanager/GetPropertiesFor(accountName=@v)?@v=",
    siteUserPrefix: "_api/web/siteusers(@v)?@v='",
    ensureUserUrl: "_api/web/ensureuser",
    accountNamePrefix: "i:0#.f|membership|",
    sendEmailPath: "/_api/SP.Utilities.Utility.SendEmail",
    formDigestRefreshInterval: 19 * 60 * 1000
  },
  config = baseConfig;

/*
    Vue installer

    Vue.use(SharePoint, "/sites/MySite/", {
        productionHosts: ["production.com"]
    })
*/
export default {
  install(Vue, baseUrl, configUpdates) {
    console.log("Base url: " + baseUrl);
    config = Object.assign(baseConfig, configUpdates);
    let sp = new SharePoint(baseUrl);
    Object.defineProperty(Vue.prototype, "$sp", { value: sp });
  }
};

export class SharePoint {
  constructor(baseUrl) {
    // properties
    this.baseUrl = baseUrl;
    this.digestValue = null;
    this.inProduction = false;

    if (!baseUrl) {
      // guess at base URL
      let paths =
          location && location.pathname
            ? location.pathname.match(/^\/\w+\/\w+\//g)
            : null,
        path = paths && paths.length ? paths[0] : null;
      this.baseUrl = path;
    }

    for (let n = 0; n < config.productionHosts.length; n++) {
      if (location.href.indexOf(config.productionHosts[n]) >= 0) {
        this.inProduction = true;
        break;
      }
    }
  }

  /* 
    Obtain SharePoint form digest value used to validate that user has authority to write data to a list in that SP site.
    If successful, calls success(digestValue)
    Also sets timer to refresh digest on a periodic basis.
  */
  getFormDigest(success, failure) {
    let me = this;
    setInterval(() => {
        me.getFormDigest().catch((error) => {return;});
        me.log("Refreshed digest value");
      }, config.formDigestRefreshInterval);
    return new Promise((resolve, reject) => {
      if (!me.inProduction) {
        me.digestValue = "dev digest value";
        if (typeof success == "function") resolve("dev digest value");
      } else {
        axios
          .post(
            me.baseUrl + "_api/contextinfo",
            {},
            {
              withCredentials: true,
              headers: {
                Accept: "application/json; odata=verbose",
                "X-HTTP-Method": "POST"
              }
            }
          )
          .then(function(response) {
            if (response && response.data) {
              if (response.data.d && response.data.d.GetContextWebInformation)
                me.digestValue =
                  response.data.d.GetContextWebInformation.FormDigestValue;
              else if (response.data.FormDigestValue)
                me.digestValue = response.data.FormDigestValue;
              resolve(me.digestValue);
            } reject("No digest provided");
          })
          .catch(function(error) {
            reject(error);
          });
      }
    });
  }

  /* True if request digest is known */
  isWriteReady() {
    return (
      (this.inProduction && this.digestValue != null) ||
      this.inProduction == false
    );
  }

  getDigest() {
    return this.digestValue;
  }

  /*
      Get data
      params = {
          baseUrl: (optional)
          path: (after baseUrl),
          success: function(listData),
          failure: function(error),
          devStaticDataUrl: ""
      }
  */
  get(params) {
    if (this.inProduction) {
      axios
        .get((params.baseUrl ? params.baseUrl : this.baseUrl) + params.path, {
          cache: false,
          withCredentials: true,
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose"
          }
        })
        .then(function(response) {
          if (
            response &&
            response.data &&
            response.data.d &&
            response.data.d.results &&
            typeof params.success == "function"
          )
            params.success.call(this, response.data.d.results);
        })
        .catch(function(error) {
          if (typeof params.failure == "function")
            params.failure.call(this, error);
        });
    } else if (params.devStaticDataUrl) {
      axios
        .get(params.devStaticDataUrl, {})
        .then(function(response) {
          if (typeof params.success == "function" && response)
            params.success.call(this, response.data);
        })
        .catch(function(error) {
          if (typeof params.failure == "function")
            params.failure.call(this, error);
        });
    } else {
      if (typeof params.failure == "function")
        params.failure.call(this, "No static data in dev");
    }
  }
  /*
      Get data from a list
      params = {
          baseUrl: <optional url>,
          listName: 'list name',
          select: 'field1,field2,field3',
          filter: "field2 eq 'value'",
          expand: 'field1',
          orderby: "field1 asc|desc",
          top: <optional number>,
          success: function(listData),
          failure: function(error),
          devStaticDataUrl: ""
      }
  */
  getList(params) {
    var q = [];
    if (params.top) q.push("$top=" + params.top);
    if (params.orderby) q.push("$orderby=" + params.orderby);
    if (params.select) q.push("$select=" + params.select);
    if (params.expand) q.push("$expand=" + params.expand);
    if (params.filter) q.push("$filter=" + params.filter);
    this.get({
      baseUrl: params.baseUrl,
      path: `${config.listPath}getbytitle('${params.listName}')/items?${q.join(
        "&"
      )}`,
      success: params.success,
      failure: params.failure,
      devStaticDataUrl: params.devStaticDataUrl
    });
  }
  /*
      Post data
      params = {
          path: ,
          data: ,
          success: function(data, itemUrl, etag),
          failure: function(error)
      }
  */
  post(params) {
    let postMe = this;
    if (params && params.path && params.data) {
      if (!this.inProduction) {
        if (config.showConsoleActivityInDev)
          this.log("Post to SharePoint: " + JSON.stringify(params.data));
        if (typeof params.success == "function")
          params.success.call(this, params.data, "dev item url");
      } else {
        let url =
          params.path.indexOf("//") > 0
            ? params.path
            : this.baseUrl + params.path;
        axios
          .post(url, params.data, {
            withCredentials: true,
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
              "X-RequestDigest": this.digestValue,
              "X-HTTP-Method": "POST"
            }
          })
          .then(function(response) {
            let data =
                response && response.data && response.data.d
                  ? response.data.d
                  : null,
              metadata = data ? data.__metadata : null,
              results = data ? data.results : null;
            if (typeof params.success == "function") {
              if (metadata)
                params.success.call(postMe, data, metadata.uri, metadata.etag);
              else if (results) params.success.call(postMe, results);
              else params.success.call(postMe, response.data);
            }
          })
          .catch(function(error) {
            if (typeof params.failure == "function")
              params.failure.call(postMe, error);
          });
      }
    }
  }
  /*
      Write data to a list, appending it as a new item
      params = {
          listName: ,
          itemData: { field1: , field2: },
          success: function(data, itemUrl),
          failure: function(error)
      }
  */
  addListItem(params) {
    if (params && params.listName && params.itemData) {
      params.data = Object.assign(
        { __metadata: { type: this.getListItemType(params.listName) } },
        params.itemData
      );
      params.path = config.listPath + `getbytitle('${params.listName}')/items`;
      this.post(params);
    }
  }
  /*
      Write data to a list, updating an existing item
      First, reloads items to obtain etag value, then posts updates using the etag value and request digest
      params = {
          listName: ,
          itemUrl:
          itemData: { field1: , field2: },
          success: function(data),
          failure: function(error)
      }
  */
  updateListItem(params) {
    let updateMe = this;
    if (params && params.itemUrl && params.itemData) {
      if (!this.inProduction) {
        this.log("Update params: " + JSON.stringify(params));
        if (typeof params.success == "function")
          params.success.call(this, params.itemData);
      } else {
        let getETagParams = {
          path: params.itemUrl,
          data: {},
          success: function(d, itemUrl, etag) {
            if (etag) {
              let updateData = Object.assign(
                {
                  __metadata: {
                    type: updateMe.getListItemType(params.listName)
                  }
                },
                params.itemData
              );
              axios
                .post(params.itemUrl, JSON.stringify(updateData), {
                  withCredentials: true,
                  headers: {
                    Accept: "application/json;odata=verbose",
                    "Content-Type": "application/json;odata=verbose",
                    "X-RequestDigest": updateMe.digestValue,
                    "X-HTTP-Method": "MERGE",
                    "If-Match": etag
                  }
                })
                .then(function(response) {
                  if (typeof params.success == "function" && response)
                    params.success.call(updateMe, response.data);
                })
                .catch(function(error) {
                  if (typeof params.failure == "function")
                    params.failure.call(updateMe, error);
                });
            }
          }
        };
        this.post(getETagParams);
      }
    }
  }

  /*
      Delete an item
      params = {
          itemUrl: <url of the item> which you can find when loading them item as item.__metadata.uri
          success: function(data),
          failure: function(error)
      }
  */
  delete(params) {
    if (params && params.itemUrl) {
      if (!this.inProduction) {
        if (config.showConsoleActivityInDev)
          this.log(
            "Delete item in SharePoint: " + JSON.stringify(params.itemUrl)
          );
        if (typeof params.success == "function") params.success.call(this);
      } else {
        axios
          .post(params.itemUrl, null, {
            withCredentials: true,
            headers: {
              "X-RequestDigest": this.digestValue,
              "IF-MATCH": "*",
              "X-HTTP-Method": "DELETE"
            }
          })
          .then(function(response) {
            if (typeof params.success == "function") params.success.call();
          })
          .catch(function(error) {
            if (typeof params.failure == "function")
              params.failure.call(this, error);
          });
      }
    }
  }

  /*
    Post a CAML query and return the response
  */
  camlQuery({ listName, queryXml, success, failure, devStaticDataUrl }) {
    let me = this;
    if (listName && queryXml) {
      this.post({
        path: `${config.listPath}getbytitle('${listName}')/getitems`,
        data: {
          query: {
            __metadata: { type: "SP.CamlQuery" },
            ViewXml: queryXml
          }
        },
        success(results) {
          if (typeof success == "function") success.call(this, results);
        },
        failure(error) {
          if (typeof failure == "function") failure.call(this, error);
        },
        devStaticDataUrl: devStaticDataUrl
      });
    }
  }

  /*
    Get ODATA-format query to retrieve matching comparisonList values
    query = getQueryFilter(["Americas", "EMEIA"], "Area")
    => (Area eq 'Americas' or Area eq 'EMEIA')
  */
  getQueryFilter(comparisonList, fieldName) {
    if (
      comparisonList &&
      typeof comparisonList == "object" &&
      comparisonList.length
    ) {
      let searchPattern = "";
      comparisonList.forEach(compare => {
        if (compare) {
          // for (let key in odataReplacements)
          //   compare = compare.replace(key, odataReplacements[key]);
          searchPattern +=
            (searchPattern.length ? " or " : "") +
            fieldName +
            " eq '" +
            encodeURIComponent(compare) +
            "'";
        }
      });
      return "(" + searchPattern + ")";
    } else return "";
  }

  /*
    Reset the array and fill it with the values of the SharePoint listItem's results array
    let countries = []
    castListValuesTo({ID:1, results:[ {Title:"USA"}, {Title:"CA"} ]}, countries, "Country")
    console.log(countries)
    => [ "USA", "CA" ]
  */
  castListValuesTo(listItem, array, valuePropName) {
    if (!valuePropName) valuePropName = "Title";
    array.splice(0);
    if (listItem && listItem.results) {
      listItem.results.forEach(item => {
        if (item && item[valuePropName]) array.push(item[valuePropName]);
      });
    }
  }
  /*
    Use this to build a results object containing the IDs of text keys
    for writing a multi-value item to SharePoint list
    based on a simple array of the text values.
    For example, if you are storing an array of country codes as text,
    and want to write it back to a SharePoint list multi-value field that
    is referencing a separate list by ID, use this method to construct the results object.
    let countryCodes = [ "FR", "CA" ];
    let fullListOfCountryCodesToIDs = ["US": 1, "CA": 2, "FR": 3];
    getIDResultsObject(fullListOfCountryCodesToIDs, countryCodes)
    => { results: [3,2] }
  */
  getIDResultsObject(fullKeyValueListOfTextToIDs, keyList) {
    if (keyList) {
      let results = [];
      keyList.forEach(key => {
        let id = fullKeyValueListOfTextToIDs.items
          ? fullKeyValueListOfTextToIDs.items[key]
          : fullKeyValueListOfTextToIDs[key];
        if (id) results.push(id);
      });
      return { results: results };
    } else return null;
  }

  /*
 Retrieve profile data and pass it to success function
 accountName is in format i:0#.f|membership|first.last@ey.com
    params = {
        accountName: i:0#.f|membership|name@domain,
        OR email: name@domain.com,
        property: <optional>,
        success: function(data),
        failure: function(error)
    }
 If property is not provided, then these fields are provided:
    success({
      PictureUrl,
      Email,
      DisplayName,
      PersonalUrl,
      Title (rank)
    })
 */
  retrievePeopleProfile(params) {
    if (this.inProduction) {
      let accountName = params.accountName;
      if (!accountName && params.email)
        accountName = config.accountNamePrefix + params.email;
      if (accountName) {
        axios
          .get(
            config.peopleManagerbaseUrl +
              "'" +
              escape(accountName) +
              "'" +
              (params.property
                ? "&$select=" + params.property
                : "&$select=" + config.profileDefaultSelect),
            {
              cache: false,
              withCredentials: true,
              headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose"
              }
            }
          )
          .then(function(response) {
            if (
              response &&
              response.data &&
              response.data.d &&
              typeof params.success == "function"
            )
              params.success.call(this, response.data.d);
          })
          .catch(function(error) {
            if (typeof params.failure == "function")
              params.failure.call(this, error);
          });
      } else if (typeof params.failure == "function")
        params.failure.call(this, "No account name provided");
    } else if (typeof params.success == "function")
      setTimeout(function() {
        params.success.call(this, {
          DisplayName: "TEST NAME",
          Email: "test@ey.com",
          PersonalUrl: "",
          Title: "Staff"
        });
      }, constDevLoadDelay);
  }
  /*
  Retrieve profile of current user
    params = {
        success: function(data),
        failure: function(error)
    }
  Promised:
    success({
      DisplayName,
      AccountName,
      Email,
      PersonalUrl,
      PicturUrl,
      Title (rank)
    })
  */
  retrieveCurrentUserProfile(params) {
    if (this.inProduction)
      axios
        .get(
          config.currentUserPropertiesPrefix + config.myProfileDefaultSelect,
          {
            cache: false,
            withCredentials: true,
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose"
            }
          }
        )
        .then(function(response) {
          if (
            response &&
            response.data &&
            response.data.d &&
            typeof params.success == "function"
          )
            params.success.call(this, response.data.d);
        })
        .catch(function(error) {
          if (typeof params.failure == "function")
            params.failure.call(this, error);
        });
    else if (typeof params.success == "function")
      setTimeout(function() {
        params.success.call(this, {
          DisplayName: "CURRENT USER"
        });
      }, constDevLoadDelay);
  }
  /*
    Check if user exists on current site. If not, add them.
    Pass their ID to the success handler.
      params = {
          accountName: i:0#.f|membership|name@domain,
          success: function(id),
          failure: function(error)
      }
  */
  ensureSiteUserId(params) {
    if (this.inProduction)
      axios
        .post(
          this.baseUrl + config.ensureUserUrl,
          {
            logonName: params.accountName
          },
          {
            cache: false,
            withCredentials: true,
            headers: {
              "X-RequestDigest": this.digestValue,
              accept: "application/json;odata=verbose"
            }
          }
        )
        .then(function(response) {
          if (
            response &&
            response.data &&
            typeof params.success == "function"
          ) {
            params.success.call(this, response.data.Id);
          }
        })
        .catch(function(error) {
          if (typeof params.failure == "function")
            params.failure.call(this, error);
        });
    else if (typeof params.success == "function")
      params.success.call(this, 1234);
  }
  /*
    Send email using the SharePoint server
      params: {
        from: "",
        to: [],
        subject: "",
        bodyHtml: "",
        success(),
        failure(error)
      }
    Promised:
      success()
  */
  sendEmail(params) {
    let mailData = {
      properties: {
        __metadata: {
          type: "SP.Utilities.EmailProperties"
        },
        From: params.from,
        To: {
          results: params.to
        },
        Subject: params.subject,
        Body: params.bodyHtml
      }
    };
    if (this.inProduction)
      axios
        .post(this.baseUrl + config.sendEmailPath, mailData, {
          withCredentials: true,
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": this.digestValue,
            "X-HTTP-Method": "POST"
          }
        })
        .then(function(response) {
          if (typeof params.success == "function") params.success.call(this);
        })
        .catch(function(error) {
          if (typeof params.failure == "function")
            params.failure.call(this, error);
        });
    else if (config.showConsoleActivityInDev) {
      this.log(`Send email to: ${params.to}`);
      this.log(`From: ${params.from}`);
      this.log(`Subject: ${params.subject}`);
      this.log(`Body: ${params.bodyHtml}`);
      if (typeof params.success == "function") params.success.call(this);
    }
  }

  log(message) {
    if (!this.inProduction) console.log(message);
  }

  getListItemType(name) {
    name = name.replace(/\s/gi, "_x0020_");
    return `SP.Data.${name[0].toUpperCase() + name.substring(1)}ListItem`;
  }
}
