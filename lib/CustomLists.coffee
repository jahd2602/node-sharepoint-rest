class CustomLists
  constructor: ()->

  getLists: (cb)->
    processRequest = (err, res, body)->
      if !body || !JSON.parse(body)
        cb({err:err})
      else
        cb(err, JSON.parse(body).value)

    config =
      headers :
        Accept: "application/json;odata=nometadata"
      strictSSL: @settings.strictSSL
      url     : "#{@url}/_api/lists"

    @request.get(config, processRequest).auth(@user, @pass, true)

    return @

  getListItemsByTitle: (title, cb)->
    return @getListItemsByTitleWithQuery(title,'',cb)

  getListItemsByTitleWithQuery: (title, query, cb)->
    processRequest = (err, res, body)->
      if !body || !JSON.parse(body)
        cb(body)
      else
        cb(err, JSON.parse(body).value)

    config =
      headers:
        Accept: "application/json;odata=nometadata"
      strictSSL: @settings.strictSSL
      url: "#{@url}/_api/web/lists/getbytitle('#{title}')/items?#{query}"

    @request.get(config, processRequest).auth(@user, @pass, true)

    return @

  getListTypeByTitle: (title, cb)->
    processRequest = (err, res, body)->
      if !body || !JSON.parse(body)
        cb("no list of title : #{title}")

      else
        cb(err, JSON.parse(body).ListItemEntityTypeFullName)

    config =
      headers :
        Accept: "application/json;odata=nometadata"
      strictSSL: @settings.strictSSL
      url     : "#{@url}/_api/web/lists/getbytitle('#{title}')?"

    @request.get(config, processRequest).auth(@user, @pass, true)

    return @

  addAttachmentToListItem: (req, cb)->
    processRequest = (err, res, body)->
      if err
        @log err

      jsonBody = JSON.parse(body)

      if jsonBody.error && jsonBody.error.code.indexOf("Microsoft.SharePoint.Client.InvalidClientQueryException") >= 0
        cb("Microsoft.SharePoint.Client.InvalidClientQueryException", null)

      if jsonBody.error && jsonBody.error.code.indexOf("Microsoft.SharePoint.SPException")
        cb("Microsoft.SharePoint.SPException", null)

      if jsonBody.error && jsonBody.error.code
        cb(jsonBody.error, null)

      else
        cb(err, JSON.parse(body))

    config =
      headers :
        "Accept": "application/json;odata=nometadata"
        "X-RequestDigest": req.context
        "content-type": "application/json;odata=nometadata"
      url     : "#{@url}/_api/web/lists/getbytitle('#{req.title}')/items(#{req.itemId})/AttachmentFiles/add(Filename='#{req.binary.fileName}')"
      strictSSL: @settings.strictSSL
      body: req.data
      binaryStringRequestBody: true
      state: "update"

    @request.post(config, processRequest).auth(@user, @pass, true)

    return @

  deleteAttachmentToListItem: (req, cb)->
    processRequest = (err, res, body)->
      if err
        @log err
      if body
        jsonBody = JSON.parse(body)

        if jsonBody.error && jsonBody.error.code.indexOf("Microsoft.SharePoint.Client.InvalidClientQueryException") >= 0
          cb("Microsoft.SharePoint.Client.InvalidClientQueryException", null)

        if jsonBody.error && jsonBody.error.code.indexOf("Microsoft.SharePoint.SPException")
          cb("Microsoft.SharePoint.SPException", null)

        if jsonBody.error && jsonBody.error.code
          cb(jsonBody.error, null)

        else
          cb(err, JSON.parse(body))
      else
        cb(err, null)
    config =
      headers :
        "Accept": "application/json;odata=nometadata"
        "X-RequestDigest": req.context
        "content-type": "application/json;odata=nometadata"
        "X-HTTP-Method": "DELETE"
        "If-Match": "*"
      url: "#{@url}/_api/web/lists/getbytitle('#{req.title}')/getItemById(#{req.itemId})/AttachmentFiles/getbyFileName('#{req.binary.fileName}')"
      strictSSL: @settings.strictSSL
      binaryStringRequestBody: true
      state: "update"

    @request.post(config, processRequest).auth(@user, @pass, true)

    return @

  addListItemByTitle: (title, item, context, cb)->
    processRequest = (err, res, body)->
      jsonBody = JSON.parse(body)

      if jsonBody.error && jsonBody.error.code.indexOf("Microsoft.SharePoint.Client.InvalidClientQueryException") >= 0
        cb("Microsoft.SharePoint.Client.InvalidClientQueryException", null)

      else if jsonBody.error && jsonBody.error.code
        cb(JSON.parse(body).error, null)

      else
        cb(err, JSON.parse(body))

    config =
      headers :
        "Accept": "application/json;odata=nometadata"
        "X-RequestDigest": context
        "content-type": "application/json;odata=nometadata"
      url: "#{@url}/_api/web/lists/getbytitle('#{title}')/items"
      strictSSL: @settings.strictSSL
      body: JSON.stringify(item)

    @request.post(config, processRequest).auth(@user, @pass, true)

    return @

  editListItemByTitle: (title, id, item, context, cb)->
    processRequest = (err, res, body)->
      try
        jsonBody = JSON.parse(body)
      catch e
        cb()
        return

      if jsonBody.error
        console.log jsonBody
        cb(jsonBody, null)
      else
        cb(err, JSON.parse(body))

    config =
      headers :
        "Accept": "application/json;odata=nometadata"
        "X-RequestDigest": context
        "content-type": "application/json;odata=nometadata"
        "X-HTTP-Method": "MERGE"
        "If-Match": "*"
      url: "#{@url}/_api/web/lists/getbytitle('#{title}')/items(#{id})"
      strictSSL: @settings.strictSSL
      body: JSON.stringify(item)

    @request.post(config, processRequest).auth(@user, @pass, true)

    return @

  deleteListItemByTitle: (title, id, context, cb)->
    processRequest = (err, res, body)->
      try
        jsonBody = JSON.parse(body)
      catch e
        cb()
        return

      if jsonBody.error
        console.log jsonBody
        cb(jsonBody, null)
      else
        cb(err, JSON.parse(body))

    config =
      headers :
        "Accept": "application/json;odata=nometadata"
        "X-RequestDigest": context
        "content-type": "application/json;odata=nometadata"
        "X-HTTP-Method": "DELETE"
        "If-Match": "*"
      url: "#{@url}/_api/web/lists/getbytitle('#{title}')/items(#{id})"
      strictSSL: @settings.strictSSL

    @request.post(config, processRequest).auth(@user, @pass, true)

    return @

  createList: (req, cb)->
    if !req.context
      cb({err: "please provide a context"})
      return

    if !req.title
      cb({err: "please provide a title"})
      return

    if !req.description
      cb({err: "please provide a description"})
      return

    context     = req.context
    title       = req.title
    description = req.description

    processRequest = (err, res, body)->
      jsonBody = JSON.parse(body)

      if jsonBody.error && jsonBody.error.code.indexOf("Microsoft.SharePoint.Client.InvalidClientQueryException") >= 0
        cb("Microsoft.SharePoint.Client.InvalidClientQueryException", null)

      else if jsonBody.error && jsonBody.error.code
        cb(JSON.parse(body).error, null)

      else
        cb(err, JSON.parse(body))

    body =
      AllowContentTypes: true
      BaseTemplate: 100
      ContentTypesEnabled: true
      Description: description
      Title: title

    config =
      headers :
        "Accept": "application/json;odata=nometadata"
        "X-RequestDigest": context
        "content-type": "application/json;odata=nometadata"
      url: "#{@url}/_api/web/lists"
      strictSSL: @settings.strictSSL
      body: JSON.stringify(body)

    @request.post(config, processRequest).auth(@user, @pass, true)

    return @

  deleteListByGUID: (req, cb)->
    if !req.context
      cb({err: "please provide a context"})
      return

    if !req.guid
      cb({err: "please provide a guid"})
      return

    if !cb
      cb({err: "please provide a callback"})
      return

    context = req.context
    guid    = req.guid

    processRequest = (err, res, body)->
      cb(err)

    config =
      headers:
        "Accept": "application/json;odata=nometadata"
        "X-RequestDigest": context
        "IF-MATCH": "*"
        "X-HTTP-Method": "DELETE"
      url: "#{@url}/_api/web/lists(guid'#{guid}')"
      strictSSL: @settings.strictSSL

    @request.post(config, processRequest).auth(@user, @pass, true)

    return @

  createColumnForListByGUID: (req, cb)->
    if !req.context
      cb({err: "please provide a context"})
      return

    if !req.title
      cb({err: "please provide a title"})
      return

    if !req.type
      cb({err: "please provide a type"})
      return

    if !req.guid
      cb({err: "please provide a guid"})
      return

    context = req.context
    title   = req.title
    type    = req.type
    guid    = req.guid

    processRequest = (err, res, body)->
      jsonBody = JSON.parse(body)
      if jsonBody.error && jsonBody.error.code
        cb(jsonBody.error, null)

      else
        cb(err, JSON.parse(body))

    body =
      Title: title
      FieldTypeKind: type
      Required: 'false'
      EnforceUniqueValues: 'false'
      StaticName: title

#    if type is 3
#      body.__metadata.RichText = "TRUE"
#      body.__metadata.RichTextMode = "FullHtml"

    config =
      headers :
        "Accept": "application/json;odata=nometadata"
        "X-RequestDigest": context
        "content-type": "application/json;odata=nometadata"
      url: "#{@url}/_api/web/lists(guid'#{guid}')/Fields"
      strictSSL: @settings.strictSSL
      body: JSON.stringify(body)

    @request.post(config, processRequest).auth(@user, @pass, true)

    return @

  getItemTypeForListName :(name) ->
    "SP.Data." + name.charAt(0).toUpperCase() + name.slice(1) + "ListItem"

  merge: (xs...) ->
    if xs?.length > 0
      @tap {}, (m)->m[k] = v for k,v of x for x in xs

  tap: (o, fn)-> fn(o); o

module.exports = CustomLists
