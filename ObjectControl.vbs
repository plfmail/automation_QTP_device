Dim IE
Function Click(object1)
	On Error Resume Next
	object1.click
	If Err.Number<>0 Then
		Gl_err=Err.Description
		Call Err_DealErr(object1)
		Call Log_logFile("Error",Gl_CaseName,Gl_err)
		Click=1
	Else
		Click=0
	End If
	On Error Goto 0
End Function

Function TypeValue(CaseCircle,stepvalue,object,value)
	On Error Resume Next
	object.type value
	caseFile=Gl_TmpCaseFileName
	stempName=CaseCircle&"&"&stepvalue
	Call Excel_WriteTmpCaseData(caseFile,stempName,"1",value)
	If Err.Number<>0 Then
		Gl_err=Err.Description
		Call Err_DealErr(object1)
		Call Log_logFile("Error",Gl_CaseName,Gl_err)
		TypeValue=1
	Else
		TypeValue=0
	End If
	On Error Goto 0
End Function

Function SetValue1(CaseCircle,stepvalue,object1,value)
	On Error Resume Next
'	tmpx=Data_GetProperty(Object,"abs_x")
'	tmpy=Data_GetProperty(Object,"abs_y")
'	tmpr=Data_GetProperty(Object,"height")
'	tmpz=Data_GetProperty(Object,"width")
'	Err_CaptureObjText Gl_err,tmpvalue,tmpx,tmpy,tmpx+tmpz,tmpy+tmpr
	object1.Set value
	caseFile=Gl_TmpCaseFileName
	stempName=CaseCircle&"&"&stepvalue
	Call Excel_WriteTmpCaseData(caseFile,stempName,"1",value)
	If Err.Number<>0 Then
		Call Err_DealErr(object1)
		Gl_err=Err.Description
		Call Log_logFile("Error",Gl_CaseName,Gl_err)
		SetValue1=1
	Else
		SetValue1=0
	End If
	On Error Goto 0
End Function

Function SelectValue(CaseCircle,stepvalue,object,value)
	On Error Resume Next
	object.select value
	caseFile=Gl_TmpCaseFileName
	stempName=CaseCircle&"&"&stepvalue
	Call Excel_WriteTmpCaseData(caseFile,stempName,"1",value)
	If Err.Number<>0 Then
		Gl_err=Err.Description
		Call Err_DealErr(object)
		Call Log_logFile("Error",Gl_CaseName,Gl_err)
		SelectValue=1
	Else
		SelectValue=0
		'MsgBox Gl_err
	End If
	On Error Goto 0
End Function

Function obj_Link(CaseCircle,stepvalue,object,value)
	On Error Resume Next
	object.Click
	If Err.Number<>0 Then
		Gl_err=Err.Description
		Call Err_DealErr(object)
		Call Log_logFile("Error",Gl_CaseName,Gl_err)
		obj_Link=1
	Else
		obj_Link=0
	End If
	'wait 3
	On Error Goto 0
End Function

Function CloseBrowser(obj)
	On Error Resume Next
	IE.Quit
	Set IE=Nothing
	On Error Goto 0
End Function

Function OpenBrowser(url)
	Dim Logicx,Logicy
	Call GetLogicxy(Logicx,Logicy)
	SystemUtil.CloseProcessByName("IEXPLORER.EXE")
	wait 2
	Set IE =CreateObject("internetexplorer application")
	IE.Visible=True
	'IE.FullScreen=true
	'IE.MenuBar=true
	'Resize IE
	hWindow=IE.HWND
	Window("hwnd ="&hWindow).Activate
	Window("hwnd ="&hWindow).Move 0,0
	Window("hwnd ="&hWindow).Resize Logicy,Logicx
	'Set Browser("MainBrowser")=IE
	wait 1
	IE.Navigate url
	OpenBrowser=0
End Function

'获取屏幕的位置，x--代表屏幕的高度，y--代表屏幕的宽度
Private Function GetLogicxy(byref x,byref y)
	strComputer="."
	Set objWMIService=GetObject("winmgmts:"_
		&"{impersonationLevel=impersonate}!\\"&strComputer&"\root\cimv2")
	Set colItems=objWMIService.ExecQuery("Select * from Win32_DesktopMonitor")
	
	For Each objItem In colItems
		x=objItem.ScreenHeight
		y=objItem.ScreenWidth
	Next
	Set colItems=Nothing
	Set objWmiService=Nothing
End Function

'	Call GetLogicxy(Logicx,Logicy)
'	MsgBox Logicx
'	MsgBox Logicy