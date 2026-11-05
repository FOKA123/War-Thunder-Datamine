"""ID-urile elementelor SAP GUI, preluate din înregistrarea VBS (Script1.vbs).

Toate ID-urile sunt centralizate aici: dacă SAP-ul (sau layout-ul) se schimbă,
se modifică doar acest fișier. Elementele marcate ``NEVERIFICAT`` nu apar în
înregistrare și trebuie confirmate cu o înregistrare scurtă (vezi README).
"""

# --- generale --------------------------------------------------------------- #
MAIN_WINDOW = "wnd[0]"
STATUS_BAR = "wnd[0]/sbar"
OKCODE = "wnd[0]/tbar[0]/okcd"
BTN_BACK = "wnd[0]/tbar[0]/btn[3]"      # F3
BTN_SAVE = "wnd[0]/tbar[0]/btn[11]"     # Ctrl+S

# --- ecranul inițial CJ02 ------------------------------------------------------ #
INITIAL_PROJECT_FIELD = "wnd[0]/usr/ctxt*PROJ-PSPID"   # NEVERIFICAT (câmpul "Definiție proiect")
INITIAL_WBS_FIELD = "wnd[0]/usr/ctxt*PRPS-POSID"       # NEVERIFICAT (se golește)
BTN_WBS_OVERVIEW = "wnd[0]/tbar[1]/btn[5]"             # din înregistrare: după salvare -> lista WBS

# --- lista elementelor WBS -------------------------------------------------- #
WBS_TABSTRIP = "wnd[0]/usr/tabsTABCJLE"
TAB_WBS_GENERAL = WBS_TABSTRIP + "/tabpLEGR"
TAB_WBS_LETR = WBS_TABSTRIP + "/tabpLETR"
TAB_WBS_ASSIGN = WBS_TABSTRIP + "/tabpLECO"
WBS_TABLE = TAB_WBS_GENERAL + "/ssubLISTE:SAPLCJWB:0902/tblSAPLCJWBTAB_902"
WBS_ASSIGN_TABLE = TAB_WBS_ASSIGN + "/ssubLISTE:SAPLCJWB:0903/tblSAPLCJWBTAB_903"

# câmp -> (nume câmp SAP, coloana din înregistrare folosită ca rezervă)
WBS_COL_LEVEL = ("PRPS-STUFE", 1)
WBS_COL_IDENT = ("RCWBS-IDENT", 2)
WBS_COL_DESC = ("PRPS-POST1", 3)
ASSIGN_COLS = {
    "STORT": ("PRPS-STORT", 7),
    "EQUNR": ("PRPS-EQUNR", 14),
    "TPLNR": ("PRPS-TPLNR", 15),
}

MENU_WBS_DETAIL = "wnd[0]/mbar/menu[4]/menu[3]"
BTN_WBS_CUSTOMER_TAB = "wnd[0]/tbar[1]/btn[19]"
CUSTOM_FIELDS = {
    "ZZCOD_BUGET": "wnd[0]/usr/ssubCUSTSCR1:SAPLXCN1:0700/ctxtPRPS-ZZCOD_BUGET",
    "ZZPARTE": "wnd[0]/usr/ssubCUSTSCR1:SAPLXCN1:0700/txtPRPS-ZZPARTE",
}

# --- activități (rețea) ------------------------------------------------------- #
TAB_ACTIVITIES = "wnd[0]/usr/tabsTABSTR199/tabpVGTS"
MENU_COPY_STD_NETWORK = "wnd[0]/mbar/menu[0]/menu[5]"
STD_NETWORK_NR = "wnd[1]/usr/ctxtRCN10-STDNR"
STD_NETWORK_ALT = "wnd[1]/usr/txtRCN10-STDAL"
POPUP_OK = "wnd[1]/tbar[0]/btn[0]"

TAB_ACT_EXTERNAL = "wnd[0]/usr/tabsTABSTRIP_2000/tabpFRML"
ACT_TABLE = TAB_ACT_EXTERNAL + "/ssubSUBSCR_2000:SAPLCOVG:2050/tblSAPLCOVGTCTRL_2050"
ACT_DELETE_BTN = TAB_ACT_EXTERNAL + "/ssubSUBSCR_2000:SAPLCOVG:2050/btnBUTTON_DELETE"
ACT_COL_VORNR = ("AFVGD-VORNR", 0)
ACT_COL_TEXT = ("AFVGD-LTXA1", 2)

# --- detaliu activitate (fereastra modală wnd[1]) ------------------------------- #
ACT_DETAIL_WINDOW = "wnd[1]"
ACT_DETAIL_TEXT = "wnd[1]/usr/subIDENTIFICATION:SAPLCONW:0100/txtAFVGM-LTXA1"
ACT_DETAIL_TAB_EXTERNAL = "wnd[1]/usr/tabsTABSTRIP_1000/tabpFRMD"
ACT_DETAIL_EKGRP = ACT_DETAIL_TAB_EXTERNAL + "/ssubSUBSCR_1000:SAPLCONW:1510/ctxtAFVGD-EKGRP"
ACT_DETAIL_AFNAM = ACT_DETAIL_TAB_EXTERNAL + "/ssubSUBSCR_1000:SAPLCONW:1510/txtAFVGD-AFNAM"
ACT_DETAIL_SERVICES_BTN = "wnd[1]/tbar[0]/btn[6]"
ACT_DETAIL_CLOSE_BTN = "wnd[1]/tbar[0]/btn[5]"

# --- specificație servicii ------------------------------------------------------ #
SRV_TABLE = "wnd[0]/usr/subSERVICE:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW"
SRV_COL_CODE = ("ESLL-SRVPOS", 2)
SRV_COL_QTY = ("ESLL-MENGE", 4)
SRV_COL_PRICE = ("ESLL-TBTWR", 6)
SRV_COL_UNLIMITED = ("ESLL-UEBTK", 8)

# --- pop-up-uri standard ------------------------------------------------------- #
POPUP_VALUE_PREFIX = "usr/sub:SAPLSPO4:0300"     # cerere de valori (ex. AMC / MBT1)
POPUP_YES = "usr/btnSPOP-VAROPTION1"
POPUP_NO = "usr/btnSPOP-VAROPTION2"
POPUP_OPTION_YES = "usr/btnSPOP-OPTION1"
POPUP_OPTION_NO = "usr/btnSPOP-OPTION2"
POPUP_ERRORLOG_EDIT = "usr/btnDY_VAROPTION2"      # salvare cu erori -> înapoi la editare
POPUP_ERRORLOG_OPT1 = "usr/btnDY_VAROPTION1"
