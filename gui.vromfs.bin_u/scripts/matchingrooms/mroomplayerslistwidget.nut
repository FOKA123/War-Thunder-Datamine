/*
 API:
 static create(config)
   config:
     scene (required) - object where need to create players lists
     teams (required) - list of teams (g_team) to show in separate columns
     room - room to gather members data (null == current SessionLobby room)
     columnsList - list of table columns to show

     onPlayerSelectCb(player) - callback on player select
     onPlayerDblClickCb(player) - callback on player double click
     onPlayerRClickCb(player) = callback on player RClick
     onWrapUpCb(obj), onWrapDownCb(obj) - callbacks for wrapUp and down.
*/


class ::gui_handlers.MRoomPlayersListWidget extends ::gui_handlers.BaseGuiHandlerWT
{
  wndType = handlerType.CUSTOM
  sceneBlkName = null
  sceneTplName = "gui/mpLobby/playersList"

  teams = null
  room = null
  columnsList = ["team", "country", "name", "status"]

  onPlayerSelectCb = null
  onPlayerDblClickCb = null
  onPlayerRClickCb = null
  onWrapUpCb = null
  onWrapDownCb = null

  playersInTeamTables = null
  focusedTeam = ::g_team.ANY
  isTablesInUpdate = false

  static TEAM_TBL_PREFIX = "players_table_"

  static function create(config)
  {
    if (!::getTblValue("teams", config) || !::check_obj(::getTblValue("scene", config)))
    {
      ::dagor.assertf(false, "cant create playersListWidget - no teams or scene")
      return null
    }
    return ::handlersManager.loadHandler(::gui_handlers.MRoomPlayersListWidget, config)
  }

  function getSceneTplView()
  {
    local view = {
      teamsAmount = teams.len()
      teams = []
    }

    local markupData = {
      tr_size = "pw, @baseTrHeight"
      columns = {
        name = { width = "fw" }
      }
    }
    local maxRows = ::SessionLobby.getMaxMembersCount(room)
    foreach(idx, team in teams)
    {
      markupData.invert <- idx == 0  && teams.len() == 2
      view.teams.append(
      {
        isFirst = idx == 0
        tableId = getTeamTableId(team)
        content = ::build_mp_table([], markupData, columnsList, maxRows)
      })
    }
    return view
  }

  function initScreen()
  {
    setFullRoomInfo()
    playersInTeamTables = {}
    focusedTeam = teams[0]
    updatePlayersTbl()
  }

  /*************************************************************************************************/
  /*************************************PUBLIC FUNCTIONS *******************************************/
  /*************************************************************************************************/

  function getFocusObj()
  {
    return getFocusedTeamTableObj()
  }

  function getSelectedPlayer()
  {
    local objTbl = getFocusedTeamTableObj()
    return objTbl && ::getTblValue(objTbl.getValue(), ::getTblValue(focusedTeam, playersInTeamTables))
  }

  function getSelectedRowPos()
  {
    local objTbl = getFocusedTeamTableObj()
    if(!objTbl)
      return null
    local rowNum = objTbl.getValue()
    if (rowNum < 0 || rowNum > objTbl.childrenCount())
      return null
    local rowObj = objTbl.getChild(rowNum)
    local topLeftCorner = rowObj.getPosRC()
    return [topLeftCorner[0], topLeftCorner[1] + rowObj.getSize()[1]]
  }



  /*************************************************************************************************/
  /************************************PRIVATE FUNCTIONS *******************************************/
  /*************************************************************************************************/

  function getTeamTableId(team)
  {
    return TEAM_TBL_PREFIX + team.id
  }

  function updatePlayersTbl()
  {
    isTablesInUpdate = true
    local playersList = ::SessionLobby.getMembersInfoList(room)
    foreach(team in teams)
      updateTeamPlayersTbl(team, playersList)
    isTablesInUpdate = false

    if (!playersInTeamTables?[focusedTeam]?.len())
      foreach(team, list in playersInTeamTables)
        if (list.len())
        {
          setFocusedTeam(team)
          return
        }

    onPlayerSelect()
  }

  function updateTeamPlayersTbl(team, playersList)
  {
    local objTbl = scene.findObject(getTeamTableId(team))
    if (!::checkObj(objTbl))
      return

    local totalRows = objTbl.childrenCount()
    local teamList = team == ::g_team.ANY ? playersList
      : ::u.filter(playersList, @(p) p.team.tointeger() == team.code)
    ::set_mp_table(objTbl, teamList, { max_rows = totalRows })
    ::update_team_css_label(objTbl)

    for(local i = 0; i < totalRows; i++)
      objTbl.getChild(i).show(i < teamList.len())
    playersInTeamTables[team] <- teamList

    //update cur value
    if (teamList.len())
    {
      local curValue = objTbl.getValue()
      local validValue = ::clamp(curValue, 0, teamList.len())
      if (curValue != validValue)
        objTbl.setValue(validValue)
    }
  }

  function getFocusedTeamTableObj()
  {
    return getObj(getTeamTableId(focusedTeam))
  }

  function updateFocusedTeamByObj(obj)
  {
    focusedTeam = ::getTblValue(::getObjIdByPrefix(obj, TEAM_TBL_PREFIX), ::g_team, focusedTeam)
  }

  function onWrapRight(obj) { wrapFocusedTeam(1) }
  function onWrapLeft(obj)  { wrapFocusedTeam(-1) }

  function wrapFocusedTeam(dir)
  {
    local newFocusedTeamIdx = ::find_in_array(teams, focusedTeam, 0) + dir
    local newFocusedTeam = ::getTblValue(newFocusedTeamIdx, teams, focusedTeam)
    if (newFocusedTeam == focusedTeam)
      return
    local newTeamObj = getObj(getTeamTableId(newFocusedTeam))
    if (!newTeamObj)
      return

    setFocusedTeam(newFocusedTeam)
  }

  function setFocusedTeam(newFocusedTeam)
  {
    focusedTeam = newFocusedTeam
    getFocusedTeamTableObj().select()
    onPlayerSelect()
  }

  function onTableClick(obj)
  {
    updateFocusedTeamByObj(obj)
    onPlayerSelect()
  }

  function onTableSelect(obj)
  {
    if (isTablesInUpdate)
      return
    updateFocusedTeamByObj(obj)
    onPlayerSelect()
  }

  function onPlayerSelect()
  {
    if (onPlayerSelectCb)
      onPlayerSelectCb(getSelectedPlayer())
  }

  function onTableDblClick()    { if (onPlayerDblClickCb) onPlayerDblClickCb(getSelectedPlayer()) }
  function onTableRClick()      { if (onPlayerRClickCb)   onPlayerRClickCb(getSelectedPlayer()) }
  function onWrapUp(obj)        { if (onWrapUpCb)         onWrapUpCb(obj) }
  function onWrapDown(obj)      { if (onWrapDownCb)       onWrapDownCb(obj) }

  function onEventLobbyMembersChanged(p)
  {
    updatePlayersTbl()
  }

  function onEventLobbyMemberInfoChanged(p)
  {
    updatePlayersTbl()
  }

  function onEventLobbySettingsChange(p)
  {
    updatePlayersTbl()
  }

  function setFullRoomInfo()
  {
    if (!room)
      return
    local fullRoom = ::g_mroom_info.get(room.roomId).getFullRoomData()
    if (fullRoom)
      room = fullRoom
  }

  function onEventMRoomInfoUpdated(p)
  {
    if (room && p.roomId == room.roomId)
    {
      setFullRoomInfo()
      updatePlayersTbl()
    }
  }
}