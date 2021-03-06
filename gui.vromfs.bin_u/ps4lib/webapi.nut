local DataBlock = require("DataBlock")
local json = require_optional("json")
local toJson = json?.to_string ?? ::save_to_json
local nativeApi = require("ps4_api.webapi")

local webApiMimeTypeBinary = "application/octet-stream"
local webApiMimeTypeImage = "image/jpeg"
local webApiMimeTypeJson = "application/json; encoding=utf-8"

local webApiMethodGet = 0
local webApiMethodPost = 1
local webApiMethodPut = 2
local webApiMethodDelete = 3

local function createRequest(api, method, path=null, params={}, data=null, forceBinary=false) {
  local request = DataBlock()
  request.apiGroup = api.group
  request.method = method
  request.path = "".concat(api.path, ((path != null) ? "/{0}".subst(path) : ""))
  request.forceBinary = forceBinary

  request.params = DataBlock()
  foreach(key,val in params) {
    if (::type(val) == "array") {
      foreach(v in val)
        request.params[key] <- v
    }
    else
      request.params[key] = val
  }

  if (::type(data) == "string")
    request.request = data
  if (::type(data) == "table")
    request.request = toJson(data)
  else if (::type(data) == "array")
    foreach(part in data)
      request.part <- part
  return request
}

local function createPart(mimeType, name, data) {
  local part = DataBlock()
  part.reqHeaders = DataBlock()
  part.reqHeaders["Content-Type"] = mimeType
  part.reqHeaders["Content-Description"] = name
  if (mimeType == webApiMimeTypeImage || mimeType == webApiMimeTypeBinary)
    part.reqHeaders["Content-Disposition"] = "attachment"

  if (mimeType == webApiMimeTypeImage)
    part.filePath = data
  else
    part.data = (::type(data) == "table") ? toJson(data) : data
  return part
}

local function makeIterable(request, pos, size) {
  // Some APIs accept either start (majority) or offset (friendlist), other param is ignored
  request.params.start = pos
  request.params.offset = request.params.start
  request.params.size = size
  request.params.limit = request.params.size
  return request
}

local function noOpCb(response, err) { /* NO OP */ }


// ------------ Session actions
local sessionApi = { group = "sdk:sessionInvitation", path = "/v1/sessions" }
local session = {
  function create(info, image, data) {
    local parts = [createPart(webApiMimeTypeJson, "session-request", info)]
    if (image != null && image.len() > 0)
      parts.append(createPart(webApiMimeTypeImage, "session-image", image))
    if (data != null && data.len() > 0)
      parts.append(createPart(webApiMimeTypeBinary, "changeable-session-data", data))
    return createRequest(sessionApi, webApiMethodPost, null, {}, parts)
  }

  function update(sessionId, sessionInfo) {
    return createRequest(sessionApi, webApiMethodPut, sessionId, {}, sessionInfo)
  }

  function join(sessionId, index=0) {
    return createRequest(sessionApi, webApiMethodPost, "".concat(sessionId,"/members"), {index=index})
  }

  function leave(sessionId) {
    return createRequest(sessionApi, webApiMethodDelete, "".concat(sessionId,"/members/me"))
  }

  function data(sessionId) {
    return createRequest(sessionApi, webApiMethodGet, "".concat(sessionId,"/changeableSessionData"))
  }

  function change(sessionId, changedata) {
    return createRequest(sessionApi, webApiMethodPut, "".concat(sessionId,"/changeableSessionData"), {}, changedata, true)
  }

  function invite(sessionId, accounts, invitedata={}) {
    if (::type(accounts) == "string")
      accounts = [accounts]
    local parts = [createPart(webApiMimeTypeJson, "invitation-request", {to=accounts})]
    if (invitedata != null && invitedata.len() > 0)
      parts.append(createPart(webApiMimeTypeBinary, "invitation-data", invitedata))
    return createRequest(sessionApi, webApiMethodPost, "".concat(sessionId,"/invitations"), {}, parts)
  }
}



// ------------ Invitation actions
local invitationApi = { group = "sdk:sessionInvitation", path = "/v1/users/me/invitations" }
local invitation = {
  function use(invitationId) {
    return createRequest(invitationApi, webApiMethodPut, invitationId, {}, {usedFlag = true})
  }

  function list() {
    return createRequest(invitationApi, webApiMethodGet, null, {fields="@default,sessionId"})
  }
}

// ------------ Profile actions
local profileApi = { group = "sdk:userProfile", path = "/v1/users/me" }
local profile = {
  function listFriends() {
    local params = { friendStatus = "friend", presenceType = "incontext" }
    return createRequest(profileApi, webApiMethodGet, "friendList", params)
  }
}

// ------------ Activity Feed actions
local feedApi = { group = "sdk:activityFeed", path = "/v1/users/me" }
local feed = {
  function post(message) {
    return createRequest(feedApi, webApiMethodPost, "feed", {}, message)
  }
}

// ----------- Commerce actions
local commerceApi = { group = "sdk:commerce" path = "/v1/users/me/container" }
local commerce = {
  function detail(products, params={}) {
    local path = ":".join(products)
    return createRequest(commerceApi, webApiMethodGet, path, params)
  }

  // listing multiple categories at once requires multiple iterators, one per category. We have one.
  function listCategory(category, params={}) {
    assert(::type(category) == "string", "[PSWA] only one category can be listed at a time")
    return createRequest(commerceApi, webApiMethodGet, category, params)
  }
}

// ---------- Entitlement actions
local entitlementsApi = { group = "sdk:entitlement", path = "/v1/users/me/entitlements"}
local entitlements = {
  function granted() {
    local params = { entitlement_type = ["service", "unified"] }
    return createRequest(entitlementsApi, webApiMethodGet, null, params)
  }
}


// ---------- Utility functions and wrappers
local function is_http_success(code) { return code != null && code >= 200 && code < 300 }

local function send(action, onResponse=noOpCb) {
  local cb = function(r) {
    local err = r?.error
    local httpErr = (!is_http_success(r?.httpStatus)) ? r.httpStatus : null
    if (httpErr && err == null)
      err = { }
    if (err && err?.code == null)
      err.code <- httpErr ? httpErr : "undefined";

    onResponse(r?.response, err)
  }

  nativeApi.send(action, cb)
}

local function fetch(action, onChunkReceived, chunkSize = 20) {
  local function onResponse(response, err) {
    // PSN responses are somewhat inconsistent, but we need proper iterators
    local entry = ((::type(response) == "array") ? response?[0] : response) || {}
    local received = (entry?.start||0) + (entry?.size||0)
    local total = entry?.total_results || entry?.totalResults || received
    if (err == null && received < total)
      send(makeIterable(action, received, chunkSize), callee())

    onChunkReceived(response, err)
  }

  send(makeIterable(action, 0, chunkSize), onResponse);
}


return {
  send = send
  fetch = fetch

  session = session
  invitation = invitation
  profile = profile
  feed = feed
  commerce = commerce
  entitlements = entitlements

  noOpCb = noOpCb
}
