root {
  background-color:t = '@modalShadeColor'
  type:t="big"

  frame {
    width:t='<<slotCountX>>(@slot_width+2@slotPaddingNoTable)+2@framePadding + 2@blockInterval<<#hasScrollBar>>+@scrollBarSize<</hasScrollBar>>'
    pos:t='0.5pw-0.5w, 1@titleLogoPlateHeight + 0.3(sh - 1@titleLogoPlateHeight - h)'
    position:t='absolute'

    <<#navBar>>
    class:t='wndNav'
    <</navBar>>
    <<^navBar>>
    class:t='wnd'
    <</navBar>>

    frame_header {
      textareaNoTab {
        id:t='header_text'
        caption:t='yes'
        text:t='<<wndTitle>>'
      }
      <<#needCloseBtn>>
      Button_close {}
      <</needCloseBtn>>
    }

    tdiv {
      width:t='pw'
      <<#filters>>
      tdiv {
        <<#isRightAlign>>
        pos:t='pw-w, 0'
        position:t='absolute'
        on_wrap_left:t='onWrapUp'
        <</isRightAlign>>
        <<^isRightAlign>>
        on_wrap_right:t='onWrapDown'
        <</isRightAlign>>
        id:t='<<id>>'
        padding:t='@blockInterval'
        behaviour:t='wrapNavigator'
        navigatorShortcuts:t='yes'
        childsActivate:t='yes'
        on_wrap_up:t='onWrapUp'
        on_wrap_down:t='onWrapDown'
        <<@boxes>>
      }
      <</filters>>
    }

    frameBlock {
      id:t = 'units_list'
      size:t= 'pw, <<slotCountY>>(@slot_height+2@slotPaddingNoTable)+2@blockInterval+2@dp'
      padding:t='@blockInterval'
      position:t='relative'
      overflow-y:t='auto'
      flow:t='h-flow'
      behaviour:t='posNavigator'
      total-input-transparent:t='yes'
      clearOnFocusLost:t='no'
      alwaysShowBorder:t='yes'
      navigatorShortcuts:t='yes'
      on_select:t='onUnitSelect'
      on_wrap_up:t='onWrapUp'
      on_wrap_down:t='onWrapDown'
      <<@unitsList>>
    }

    <<#navBar>>
    navBar {
      include "gui/commonParts/navBar"
    }
    <</navBar>>
  }
}
