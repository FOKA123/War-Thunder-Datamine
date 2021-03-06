local enums = ::require("sqStdlibs/helpers/enums.nut")

::g_ww_map_info_type <- {
  types = []
  cache = {
    byIndex = {}
  }
}

::g_ww_map_info_type.template <- {
  getMainBlockHandler = function(placeObj, side = ::SIDE_NONE) { return null }
}

enums.addTypesByGlobalName("g_ww_map_info_type", {
  UNKNOWN = {
    index = -1
  }

  OBJECTIVE = {
    index = 0
    getMainBlockHandler = function(placeObj, side = null)
    {
      return ::handlersManager.loadHandler(::gui_handlers.wwObjective, {
        scene = placeObj,
        side = side || ::ww_get_player_side()
        restrictShownObjectives = true
      })
    }
  }

  LOG = {
    index = 1
    getMainBlockHandler = function(placeObj, side = null)
    {
      return ::handlersManager.loadHandler(::gui_handlers.WwOperationLog, {
        scene = placeObj,
        side = side || ::ww_get_player_side()
      })
    }
  }
}, null, "name")

g_ww_map_info_type.getTypeByIndex <- function getTypeByIndex(index)
{
  return enums.getCachedType("index", index, ::g_ww_map_info_type.cache.byIndex, ::g_ww_map_info_type, ::g_ww_map_info_type.UNKNOWN)
}
