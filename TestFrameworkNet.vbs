Dim Action_ :Set Action_ = CreateObject("Wscript.Shell")
Dim Obj_, Ver, Respond

Dim Regstry :Regstry = Array( _
	"1 - HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\1.0.3705\Version", _
	"1.1 - HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\1.1.4322\Version", _
	"2 - HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727\Version", _
	"3 - HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.0\Version", _
	"3.5 - HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5\Version", _ 
	"4 - HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\Version")
On Error Resume Next 
	For Each Obj_ In Regstry
		Ver = Split(Obj_," - ")
		If IsNull(Action_.RegRead(Ver(1))) Then
			Respond = Respond &  "Not Installed: " & Ver(0) & vbCrLf 
		Else
			Respond = Respond &  "Currently Installed: " & Ver(0) & " - " & "[" & Action_.RegRead(Ver(1)) & "]" & vbCrLf 
		End If
	Next 

MsgBox Respond, 64,".Net Framework Information"