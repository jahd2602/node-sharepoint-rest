class SharePoint
  constructor: (@settings)->
    if !@settings
      throw new Error("settings object is required for instance creation")
    else
      if !@settings.strictSSL
        @settings.strictSSL = false

      @request  = require 'request'

      @user = @settings.username || undefined
      @pass = @settings.password || undefined
      @url  = @settings.url      || undefined

      if typeof @url is "undefined" || typeof @user is "undefined" || typeof @pass is "undefined"
        throw new Error("settings object requires username, password, and url for instance creation")

      @setSiteUrl @url

  log: (msg)->
    console.log msg

  setSiteUrl: (@url)->
    @log 'setting site url to: ' + @url
    return @

  getLists: (cb)->
    processRequest = (err, res, body)->
      if !body || !JSON.parse(body).d
        cb({err:err})
      else
        cb(err, JSON.parse(body).d.results)

    config =
      headers :
        Accept: "application/json;odata=verbose"
      strictSSL: @settings.strictSSL
      url     : "#{@url}/_api/lists"

    @request.get(config, processRequest).auth(@user, @pass, true)

    return @

  getListItemsByTitle: (title, cb)->
    processRequest = (err, res, body)->
      if !body || !JSON.parse(body).d
        cb("no list of title : #{title}")
      else
        cb(err, JSON.parse(body).d.results)

    config =
      headers:
        Accept: "application/json;odata=verbose"
      strictSSL: @settings.strictSSL
      url: "#{@url}/_api/web/lists/getbytitle('#{title}')/items"

    @request.get(config, processRequest).auth(@user, @pass, true)

    return @

  getListTypeByTitle: (title, cb)->
    processRequest = (err, res, body)->
      if !body || !JSON.parse(body).d
        cb("no list of title : #{title}")

      else
        cb(err, JSON.parse(body).d.ListItemEntityTypeFullName)

    config =
      headers :
        Accept: "application/json;odata=verbose"
      strictSSL: @settings.strictSSL
      url     : "#{@url}/_api/web/lists/getbytitle('#{title}')?"

    @request.get(config, processRequest).auth(@user, @pass, true)

    return @

  getContext: (app, cb)->
    processRequest = (err, res, body)->
      if !body || !JSON.parse(body).d
        console.log "no list of title: #{app}"
      else
        cb(err, JSON.parse(body).d.GetContextWebInformation.FormDigestValue)

    config =
      headers :
        Accept: "application/json;odata=verbose"
      strictSSL: @settings.strictSSL
      url     : "#{@url}/#{app}/_api/contextinfo"

    @request.post(config, processRequest).auth(@user, @pass, true)

    return @

  addAttachment: (req, cb)->
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
        cb(err, JSON.parse(body).d)

    config =
      headers :
        "Accept": "application/json;odata=verbose"
        "X-RequestDigest": req.context
        "content-type": "application/json;odata=verbose"
      url     : "#{@url}/_api/web/lists/getbytitle('#{req.title}')/items(#{req.itemId})/AttachmentFiles/add(Filename='#{req.binary.fileName}')"
      strictSSL: @settings.strictSSL
      body: req.data
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
        cb(err, JSON.parse(body).d)

    config =
      headers :
        "Accept": "application/json;odata=verbose"
        "X-RequestDigest": context
        "content-type": "application/json;odata=verbose"
      url: "#{@url}/_api/web/lists/getbytitle('#{title}')/items"
      strictSSL: @settings.strictSSL
      body: JSON.stringify(item)

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
        cb(err, JSON.parse(body).d)

    body =
      __metadata:
        type: 'SP.List'
      AllowContentTypes: true
      BaseTemplate: 100
      ContentTypesEnabled: true
      Description: description
      Title: title

    config =
      headers :
        "Accept": "application/json;odata=verbose"
        "X-RequestDigest": context
        "content-type": "application/json;odata=verbose"
      url: "#{@url}/_api/web/lists"
      strictSSL: @settings.strictSSL
      body: JSON.stringify(body)

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
        cb(err, JSON.parse(body).d)

    body =
      __metadata:
        type: 'SP.Field'
      Title: title
      FieldTypeKind: type
      Required: 'false'
      EnforceUniqueValues: 'false'
      StaticName: title

    config =
      headers :
        "Accept": "application/json;odata=verbose"
        "X-RequestDigest": context
        "content-type": "application/json;odata=verbose"
      url: "#{@url}/_api/web/lists(guid'#{guid}')/Fields"
      strictSSL: @settings.strictSSL
      body: JSON.stringify(body)

    @request.post(config, processRequest).auth(@user, @pass, true)

    return @

module.exports = SharePoint