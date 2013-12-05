SuperClass   = require './SuperClass'
UserProfiles = require './UserProfiles'
CustomLists  = require './CustomLists'

class SharePoint extends SuperClass
  @include UserProfiles
  @include CustomLists

  constructor: (@settings)->
    UserProfiles.call @
    CustomLists.call @

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

module.exports = SharePoint