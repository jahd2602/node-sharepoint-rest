node-sharepoint-rest
=============

On Premise SharePoint 2013 REST API wrapper.  Requires basic auth to be enabled on SharePoint's IIS server.


### Using node-sharepoint-rest

    $ npm install node-sharepoint-rest

With it now installed in your project:

    settings =
        user      : "node"
        pass      : "password"
        url       : "https://sharepoint/subsite"
        strictSSL : false

    SharePoint = require 'node-sharepoint-rest'
    
    sharePoint = new SharePoint(settings)

----

#### GET API
##### getLists
is a prototype function of the SharePoint class that uses the require module and basic auth to
communicate with On Premise SharePoint instances.  It takes a callback (err, data), where data is the RESTful
results from SharePoint.

    sharePoint.getLists (err, data)->
      if err
        console.log err
      else
        console.log data[id].Title for id of data


The data returned looks something like this.  The "..." indicates a continuation:
```javascript
[
  {
    Title: 'list1',
    Created: '2013-06-18T13:51:35Z',
    Id: '12345ba1-65cb-1234-1234642ds',
    ...,
    __metadata: {
       ...
    }
  },
  {
    Title: 'list2',
    ...
  }
  ...
]
```

----

##### getListItemsByTitle
is a prototype function of the SharePoint class that uses the require module and basic auth to
grab an array of list items.  It takes a list title string and a callback (err, data),
where data is the RESTful results from SharePoint.

    sharePoint.getListItemsByTitle 'customList', (err, data)->
      if err
        console.log err
      else
        console.log data[id].Title for id of data


The data returned looks something like this.  The "..." indicates a continuation:
```javascript
[
  {
    Title: 'First Custom Item',
    Created: '2013-06-18T13:51:35Z',
    Id: 1,
    GUID: '12345ba1-65cb-1234-1234642ds',
    ...,
    __metadata: {
       ...
    }
  },
  {
    Title: 'Second Custom Item',
    ...
  }
  ...
]
```

----

The following prototype functions (getListTypeByTitle and getContext) return necessary data required to use
the wrapped POST requests for adding list items and adding list item attachments.

##### getListTypeByTitle
is a prototype function for getting the internal string name of a SharePoint list item type.
Think of it as a schema name.  It takes a list title string and a callback (err, data), where data
is the RESTful results from SharePoint in string format.

    sharePoint.getListTypeByTitle 'customList', (err, type)->
      if err
        console.log err
      else
        console.log type

    >> SP.Data.customListListItem

##### getContext
is a prototype function for getting the context of a SharePoint app.  For example, you can use
this fn to get the context string from a custom list, which is required for POSTing new items to that list.

    sharePoint.getContext 'appName', (err, context)->
      if err
        console.log err
      else
        console.log context

    >> 0x7F21F8A4C70FB50B05A7DE7DD23334D06D4BC45FAAZZFFDDDDD32D125B556A278AFF84B948BE99AC0045DFCC3B25F8D01F24A94B8DF10C36CE1C1B1,28 Aug 2013 16:00:07 -0000

----

#### POST API
In order to use these POST functions, you must obtain the context and list item type, first, either by storing
them before posting the request, or inline through callbacks as shown below.

##### addListItemByTitle
is a prototype function for adding an item to a custom list.  This fn takes a list title string, an item object
(containing the list type in its __metadata), the app context, and a callback (err, newItem), where newItem is
the new item from SharePoint.

    list = 'customList'
    sharePoint.getListTypeByTitle list, (err, type)->
      if err
        console.log err
        return

      sharePoint.getContext list, (err, context)->
        if err
          console.log err
          return

        item =
          __metadata:
            type: type
          Title: "My New Item " + Math.random()

        sharePoint.addListItemByTitle list, item, context, (err, newItem)->
          if err
            console.log err
          else
            console.log newItem

The data returned looks something like this.  The "..." indicates a continuation:
```javascript
{
  Title: 'My New Item 0.78239873624978',
  Created: '2013-06-18T13:51:35Z',
  Id: 1,
  GUID: '12345ba1-65cb-1234-1234642ds',
  ...,
  __metadata: {
     ...
  }
}
```

----

##### addAttachment
is a prototype function for adding a binary attachment to a custom list item.  This fn takes a config object
and a callback (err, data), where data is meta from the item.

    fs = require 'fs'

    binary =
      fileName: "test.txt"

    data = fs.readFileSync binary.fileName, { encoding: null }

    list = 'customList'
    sharePoint.getContext list, (err, context)->
      if err
        console.log err
        return

      req =
        title   : list
        itemId  : 1
        context : context
        data    : data
        binary  : binary

      sharePoint.addAttachment req, (err, newItem)->
        if err
          console.log err
        else
          console.log newItem
