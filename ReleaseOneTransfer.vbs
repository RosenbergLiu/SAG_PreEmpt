If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]/tbar[0]/okcd").text = "vl10d"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB5").select
session.findById("wnd[0]/usr/ctxtST_LEDAT-LOW").text = ""
session.findById("wnd[0]/usr/ctxtST_LEDAT-HIGH").text = ""
session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB5/ssub%_SUBSCREEN_ORDER_CRITERIA:RVV50R10C:1030/ctxtST_EBELN-LOW").text = WScript.Arguments(0)
session.findById("wnd[0]/usr/ctxtST_LEDAT-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtST_LEDAT-HIGH").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[5]").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/usr/lbl[73,6]").setFocus
session.findById("wnd[0]/usr/lbl[73,6]").caretPosition = 6
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV50A:1502/ctxtLIKP-VBELN").caretPosition = 10
