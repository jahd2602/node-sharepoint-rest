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
  - [getContext(app, cb)](#getcontext)
  - [getPropertiesForAccountName(accountName, cb)](#getpropertiesforaccountname)
  - [getLists(cb)](#getlists)
  - [getListItemsByTitle(title, cb)](#getlistitemsbytitle)
  
#### POST API
  - [createList(req, cb)](#createlist)
  - [createColumnForListByGUID(req, cb)](#createcolumnforlistbyguid)
  - [addListItemByTitle(listTitle, cb)](#addlistitembytitle)
  - [addAttachmentToListItem(req, cb)](#addattachment)

##### getContext
is a prototype function of the SharePoint class for getting the context of a SharePoint app.  For example, you can use
this fn to get the context string from a custom list, which is required for POSTing new items to that list.  You can also
provide an empty set '' for the appName to get the context of the site root.

    sharePoint.getContext 'appName', (err, context)->
      if err
        console.log err
      else
        console.log context

    >> 0x7F21F8A4C70FB50B05A7DE7DD23334D06D4BC45FAAZZFFDDDDD32D125B556A278AFF84B948BE99AC0045DFCC3B25F8D01F24A94B8DF10C36CE1C1B1,28 Aug 2013 16:00:07 -0000 
    
###### [Back to top](#node-sharepoint-rest)
----
  
##### <a id="getPropertiesForAccountName">getPropertiesForAccountName</a>
is a prototype function of the UserProfile class inherited by the SharePoint class.
It takes an account name string of "&lt;domain&gt;\\username" and a callback (err, data), where data is
the RESTful results from SharePoint.

    sharePoint.getPropertiesForAccountName "domain\\admin", (err, profile)->
      util = require 'util'
      console.log util.inspect(profile, {depth:4})


The data returned looks something like this.  The "..." indicates a continuation:
```javascript
{
  AccountName: 'domain\\admin',
  DirectReports: { results: [] },
  DisplayName: 'admin',
  Email: 'admin@domain.net',
  ExtendedManagers: { results: [] },
  ExtendedReports: { results: [ 'domain\\admin' ] },
  IsFollowed: true,
  LatestPost: "Hello #sharepoint citizens!",
  Peers: { results: [] },
  PersonalUrl: 'https://domain.net/personal/admin/',
  PictureUrl: 'https://my.domain.net/User%20Photos/Profile%20Pictures/admin_MThumb.jpg',
  Title: "SharePoint Admin",
  UserProfileProperties:
  {
    results: [
      {
         __metadata: { ... },
         Key: 'UserProfile_GUID',
         Value: '12341234-3934-2934-h38h-2398482348',
         ValueType: 'Edm.String'
      },
      {
         __metadata: { ... },
          Key: 'SID',
          Value: 's-1-5-20-23423423-123412342-12342341-1000',
          ValueType: 'Edm.String'
      },
      ...
    ]
  }
  ...,
  __metadata: {
     ...
  }
}
```
###### [Back to top](#node-sharepoint-rest)
----

##### <a id="getLists">getLists</a>
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
###### [Back to top](#node-sharepoint-rest)
----

##### <a id="getListItemsByTitle">getListItemsByTitle</a> 
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
###### [Back to top](#node-sharepoint-rest)
----

The following prototype functions (getListTypeByTitle and getContext) return necessary data required to use
the wrapped POST requests for adding list items and adding list item attachments.

##### <a id="getListTypeByTitle">getListTypeByTitle</a> 
is a prototype function for getting the internal string name of a SharePoint list item type.
Think of it as a schema name.  It takes a list title string and a callback (err, data), where data
is the RESTful results from SharePoint in string format.

    sharePoint.getListTypeByTitle 'customList', (err, type)->
      if err
        console.log err
      else
        console.log type

    >> SP.Data.customListListItem

#### POST API
In order to use these POST functions, you must obtain the context and list item type, first, either by storing
them before posting the request, or inline through callbacks as shown below.

##### createList
is a prototype function for creating a custom list.  You must have a recent context token first.  The example shows
a callback pattern starting with the getContext fn.  The createList fn takes a req object which must contain a
list title string, the app context, and a callback (err, list), where list is the new list from SharePoint.  As an
interesting side note:  any context will work here, even the context of the site itself (empty string) as seen
in the example below.

    # get our context from the site url (leave app parameter as empty string)
    sharePoint.getContext '', (err, context)->
      console.log err || context

      req =
        context: context
        title  : 'testing101'
        description: "This is my new custom list"

      sharePoint.createList req, (err, list)->
        console.log err || list

The data returned looks something like this.  Id will be used when calling list related fns that require GUID.
The "..." indicates a continuation:
```javascript
{
  Title: testing101',
  Created: '2013-06-18T13:51:35Z',
  Id: '12345ba1-65cb-1234-1234642ds',

  ...,
  __metadata: {
     ...
  }
}
```
###### [Back to top](#node-sharepoint-rest)
----

##### deleteListByGUID
is a prototype function for deleting a custom list.  The deleteListByGUID fn takes a req object which must contain a
list GUID, the app context, and a callback (err, meta), where meta is the response for the operation.

    # get our context from the site url (leave app parameter as empty string)
    sharePoint.getContext 'app', (err, context)->
      console.log err || context

      req =
        context: context
        guid: '47c61990-d8ae-4f73-bae7-8fdaeecb4e98'

      sharePoint.deleteListByGUID req, (err, meta)->
        console.log err || meta

The data returned looks something like this.
The "..." indicates a continuation:
```javascript
{
  Title: testing101',
  Created: '2013-06-18T13:51:35Z',
  Id: '12345ba1-65cb-1234-1234642ds',

  ...,
  __metadata: {
     ...
  }
}
```
###### [Back to top](#node-sharepoint-rest)
----

##### createColumnForListByGUID
is a prototype function for creating a custom list column.  You must have a recent context token from your list app
first.  The example shows a callback pattern starting with the getContext fn.  The
fn takes a req object which must contain a column title string, the app context,
the column type (see: http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.fieldtype.aspx),
the list GUID, and a callback (err, data), where data is the new column meta-data from SharePoint.

    sharePoint.getContext 'testing101', (err, context)->
      console.log err || context

      req =
        context: context
        title  : 'text'
        type   : 3
        guid   : '47c61990-d8ae-4f73-bae7-8fdaeecb4e98'

      sharePoint.createColumnForListByGUID req, (err, data)->
        console.log err || data

The data returned looks something like this.  Id will be used when calling list related fns that require GUID.
The "..." indicates a continuation:
```javascript
{
  Title: text',
  Created: '2013-06-18T13:51:35Z',
  Id: '12345ba1-65cb-1234-1234642ds',
  Description: '',
  FieldTypeKind: 3,
  ...,
  __metadata: {
     ...
  }
}
```
###### [Back to top](#node-sharepoint-rest)
----

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
###### [Back to top](#node-sharepoint-rest)
----

##### addAttachmentToListItem
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

      sharePoint.addAttachmentToListItem req, (err, newItem)->
        if err
          console.log err
        else
          console.log newItem
