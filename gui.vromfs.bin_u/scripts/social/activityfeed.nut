local psn = require("ps4Lib/webApi.nut")
local statsd = require("statsd")

::FACEBOOK_POST_WALL_MESSAGE <- false

::ps4_activityFeed_requestsTable <- {
  player = "$USER_NAME_OR_ID",
  count = "$STORY_COUNT",
  onlineUserId = "$ONLINE_ID",
  productName = "$PRODUCT_NAME",
  titleName = "$TITLE_NAME",
  fiveStarValue = "$FIVE_STAR_VALUE",
  sourceCount = "$SOURCE_COUNT"
}

::prepareMessageForWallPostAndSend <- function prepareMessageForWallPostAndSend(config, customFeedParams = {}, reciever = bit_activity.NONE)
{
  local copy_config = clone config
  local copy_customFeedParams = clone customFeedParams
  if (reciever & bit_activity.PS4_ACTIVITY_FEED)
    ::ps4PostActivityFeed(copy_config, copy_customFeedParams)
  if (reciever & bit_activity.FACEBOOK)
    ::facebookPostActivityFeed(copy_config, copy_customFeedParams)
}

// specialization getters below expect valid data, validated by the caller
::getActivityFeedImageByParam <- function getActivityFeedImageByParam(feed, imagesConfig)
{
  local config = imagesConfig.other?[feed.blkParamName]

  if (u.isString(config))
    return imagesConfig.mainPart + config

  if (u.isDataBlock(config) && config?.name)
  {
    local url = imagesConfig.mainPart + config.name + feed.imgSuffix
    if (config?.variations)
      url += ::format("_%.2d", ::math.rnd() % config.variations + 1)
    return url
  }

  ::dagor.debug("getActivityFeedImagesByParam: no image name in '"+feed.blkParamName)
  debugTableData(config)
  return ""
}

::getActivityFeedImageByCountry <- function getActivityFeedImageByCountry(feed, imagesConfig)
{
  local aircraft = ::getAircraftByName(feed.unitNameId)
  local esUnitType = ::get_es_unit_type(aircraft)
  local unit = ::getUnitTypeText(esUnitType)
  local country = feed.country

  local variants = imagesConfig?[country]?[unit]
  if (u.isDataBlock(variants))
    return imagesConfig.mainPart + variants.getParamValue(::math.rnd() % variants.paramCount())

  ::dagor.debug("getActivityFeedImagesByCountry: no config for '"+country+"/"+unit+" ("+feed.unitNameId+")")
  debugTableData(imagesConfig)
  return ""
}

::getActivityFeedImages <- function getActivityFeedImages(feed)
{
  local guiBlk = ::configs.GUI.get()
  local imagesConfig = guiBlk?.activity_feed_image_url
  if (u.isEmpty(imagesConfig))
  {
    ::dagor.debug("getActivityFeedImages: empty or missing activity_feed_image_url block in gui.blk")
    return null
  }

  local feedUrl = imagesConfig?.mainPart
  local imgExt = imagesConfig?.fileExtension
  if (!feedUrl || !imgExt)
  {
    ::dagor.debug("getActivityFeedImages: invalid feed config, url base '"+feedUrl+"', image extension '"+imgExt)
    debugTableData(imagesConfig)
    return null
  }

  local logo = imagesConfig?.logoEnd || ""
  local big = imagesConfig?.bigLogoEnd || ""
  local ext = imagesConfig.fileExtension
  local url = ""
  if (!u.isEmpty(feed?.blkParamName) && !u.isEmpty(imagesConfig?.other))
    url = getActivityFeedImageByParam(feed, imagesConfig)
  else if (!u.isEmpty(feed?.country) && !u.isEmpty(feed?.unitNameId))
    url = getActivityFeedImageByCountry(feed, imagesConfig)

  if (!u.isEmpty(url))
    return {
      small = url + (feed?.shouldForceLogo ? logo : "") + ext
      large = url + big + ext
    }

  ::dagor.debug("getActivityFeedImages: could not select method to build image URLs from gui.blk and feed config")
  debugTableData(feed)
  return null
}

//---------------- <Facebook> --------------------------
::facebookPostActivityFeed <- function facebookPostActivityFeed(config, customFeedParams)
{
  if (!::has_feature("FacebookWallPost"))
    return

  if ("requireLocalization" in customFeedParams)
    foreach(name in customFeedParams.requireLocalization)
      customFeedParams[name] <- ::loc(customFeedParams[name])

  ::FACEBOOK_POST_WALL_MESSAGE = false
  local locId = ::getTblValue("locId", config, "")
  if (locId == "")
  {
    ::dagor.debug("facebookPostActivityFeed, Not found locId in config")
    ::debugTableData(config)
    return
  }

  customFeedParams.player <- ::my_user_name
  local message = ::loc("activityFeed/" + locId, customFeedParams)
  local link = ::getTblValue("link", customFeedParams, "")
  local backgroundPost = ::getTblValue("backgroundPost", config, false)
  ::make_facebook_login_and_do((@(link, message, backgroundPost) function() {
                 if (!backgroundPost)
                  ::scene_msg_box("facebook_login", null, ::loc("facebook/uploading"), null, null)
                 ::facebook_post_link(link, message)
               })(link, message, backgroundPost), this)
}
//------------------ </Facebook> --------------------------------

//----------------- <PlayStation> -------------------------------
::ps4PostActivityFeed <- function ps4PostActivityFeed(config, customFeedParams)
{
  if (!::is_platform_ps4 || !::has_feature("ActivityFeedPs4"))
    return

  local sendStat = function(tags) {
    local qualifiedNameParts = split(::getEnumValName("ps4_activity_feed", config.subType, true), ".")
    tags["type"] <- qualifiedNameParts[1]
    statsd.send_counter("sq.activityfeed", 1, tags)
  }

  local locId = ::getTblValue("locId", config, "")
  if (locId == "" && u.isEmpty(customFeedParams?.captions))
  {
    sendStat({action = "abort", reason = "no_loc_id"})
    ::dagor.debug("ps4PostActivityFeed, Not found locId in config")
    ::debugTableData(config)
    return
  }

  local localizedKeyWords = {}
  if ("requireLocalization" in customFeedParams)
    foreach(name in customFeedParams.requireLocalization)
      localizedKeyWords[name] <- ::get_localized_text_with_abbreviation(customFeedParams[name])

  local activityFeed_config = ::combine_tables(::ps4_activityFeed_requestsTable, customFeedParams)

  local getFilledFeedTextByLang = function(key) {
    local captions = {}
    local localizedTable = ::get_localized_text_with_abbreviation(key)

    foreach(lang, string in localizedTable)
    {
      local localizationTable = {}
      foreach(name, value in activityFeed_config)
        localizationTable[name] <- localizedKeyWords?[name][lang] ?? value

      captions[lang] <- string.subst(localizationTable)
    }

    return captions
  }

  local images = ::getActivityFeedImages(customFeedParams)
  local largeImage = customFeedParams?.images?.large || images?.large
  local smallImage = customFeedParams?.images?.small || images?.small

  local body = {
    captions = customFeedParams?.captions || getFilledFeedTextByLang("activityFeed/" + locId)
    condensedCaptions = customFeedParams?.condensedCaptions || getFilledFeedTextByLang("activityFeed/" + locId + "/condensed")
    storyType = "IN_GAME_POST"
    subType = config?.subType || 0
    targets = [{accountId=::ps4_get_account_id(), type="ONLINE_ID"}]
  }
  if (largeImage)
    body.targets.append({meta=largeImage, type="LARGE_IMAGE_URL"})
  if (smallImage)
    body.targets.append({meta=smallImage, type="SMALL_IMAGE_URL", aspectRatio="2.08:1"})

  psn.send(psn.feed.post(body), function(_, err) {
      local tags = { action = "post", result = err ? "fail" : "success" }
      if (err != null)
        tags["reason"] <- err.code.tostring()
      sendStat(tags)
  })
}
//----------------------- </PlayStation> --------------------------
