class UserProfiles
  constructor: ()->

  getPropertiesForAccountName: (accountName, cb)->
    processRequest = (err, res, body)->
      if !body || !JSON.parse(body).d
        cb("no account of : #{accountName}")
      else
        cb(err, JSON.parse(body).d)

    config =
      headers:
        Accept: "application/json;odata=verbose"
      strictSSL: @settings.strictSSL
      url: "#{@url}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='#{accountName}'"

    @request.get(config, processRequest).auth(@user, @pass, true)

    return @

module.exports = UserProfiles
