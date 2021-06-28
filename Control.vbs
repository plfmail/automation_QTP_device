
'TestCase根据传入的测试用例文件名称（Excel格式），解析该Excel文件
'Excel的第一列是步骤数，第二列是操作Operation，第三列是Web对象的描述信息，第四列是该操作需要的数据
Function TestCase(fileName,objectFile,caseName,glLevel)
On Error Resume Next
Dim rowCount,colcount,ret
	InitDR(objectfile)
	
	LogInit("c:\temp\AutoTest.log") 
	Call Excel_CreateCaseTmpDataFile(caseName)
	Gl_CaseName = caseName
	Gl_TmpCaseFileName = "c:\temp\"&caseName&".xls"
	Gl_LogLevel = glLevel
	Word_Init
	Call Log_logFile("Debug","TestCase",":CaseName="&caseName&":Gl_TmpCaseFileName="&Gl_TmpCaseFileName)
	OpenExcel(fileName)
	OpenSheet(caseName)
	rowcount = GetRows
	colrount = GetColumns
	Call Log_LogFile("Debug","TestCase",":CaseRows="&rowcount&":caseColumns="&colrount)
'	MsgBox rowcount&colrount
	CaseCircle = 1
	
	For i = 2 To rowcount
		Gl_ErrBitmapName=""
		stepvalue=GetValue(i,1)
		Operation=GetValue(i,2)
		Data=GetValue(i,4)
		objectValue=GetValue(i,3)
		
		Call Log_LogFile("Debug","TestCase",":stepvalue="&stepvalue&":Operation="&Operation&":objectValue="&objectValue&":Data="&Data)
		If StrComp(Left(Operation,5),"Check")=0 Then
		'MsgBox "before"
			resultvalue=GetValue(i,7)
			ret=OperationSelect(CaseCircle,stepvalue,Operation,resultvalue,objectValue)	
		Else 
			ret=OperationSelect(CaseCircle,stepvalue,Operation,Data,objectValue)
		End If
		Word_AddHeader "step",stepvalue
		Word_AddWord "Operation",Operation,0
		Word_AddWord "Data",Data,0
		If ret=0 Then
			Word_AddWord "Result","Success",0
		Else
			Word_AddWord "Result","Wrong"&vbCrLf&"Description="&Gl_err,1
			If Gl_ErrBitmapName<>"" Then
				Word_Addbitmap Gl_ErrBitmapName
			End If
			
			Exit For
		End If
			Word_AddEnter
		Next
		Word_SaveWord("c:\temp\"&caseName&"Test Result-"&RandomNumber(1,10000)&".doc")
		'Set wordTest=Nothing
		Word_Terminate
		Terminate
	On Error Goto 0
End Function

Function OperationSelect(CaseCircle,stepvalue,oper,Data,objectValue)

	Select Case oper
	Case "OpenBrowser"
		'MsgBox objectValue
		'ObjectSet(objectValue)
		ret=OpenBrowser(Data)
		wait 2
	Case "SetValue"
		'MsgBox "123"
		'Object(objectvalue).set Data
		Set obj=ObjectSet(objectvalue)
		ret=SetValue1(CaseCircle,stepvalue,obj,Data)
	Case "Click"
		Set obj=ObjectSet(objectvalue)
		ret=Click(obj)
		'wait 2
	Case "CloseBrowser"
		Set obj=ObjectSet(objectValue)
		ret=CloseBrowser(obj)
	Case "CheckPage"
		Set obj=ObjectSet(objectValue)
		tmpvalue=Data_checkPage(obj,Data)
		If tmpvalue=1 Then
			Call Log_logFile("Debug","TestCase","Check Page Correct")
		Else
			Call Log_logFile("Debug","TestCase","Result InCorrect")
			Call Err_CapturePage("Correct Page Title:"&Data)
		End If
	Case "Link"
		Set obj=ObjectSet(objectValue)
		ret=obj_Link(CaseCircle,stepvalue,obj,Data)
	Case "SelectList"
		Set obj=ObjectSet(objectValue)
		If StrComp(Data,"Radom")=0 Then
			tmpvalue=Data_ListSelect(obj,0,"1")
		Else
			tmpvalue=Data_ListSelect(obj,Data,"3")
		End If
		'tempvalue="12345"
		'MsgBox tmpvalue
		If StrComp(tmpvalue,"NotFind")<>0 Then
			ret=SelectValue(CaseCircle,stepvalue,obj,tmpvalue)
		End If
	Case "CheckPageText"
		'MsgBox "CheckPageText"
		Set obj=ObjectSet(objectValue)
		tmpvalue=Data_GetPage(obj)
		'MsgBox tmpvalue
		
		If Data_JSearch(Data,tmpvalue,1)>0 Then
			Call Err_CapturePage("Result Correct:"&Data)
			Call Log_logFile("Debug",Gl_CaseName,"Check Page Text Correct:"&Data)
		Else
			Call Err_CapturePage("Result InCorrect. Cant't find Text="&Data)
			Call Log_logFile("Debug",Gl_CaseName,"Cant't find Text="&Data)
		End If
	End Select
	
	If ret=1 Then
		OperationSelect=1
	Else 
		OperationSelect=0
	End If

End Function

Function RecoveryErr(Object,Method,Arguments,retVal)
	tmpx=Data_GetProperty(Object,"abs_x")
	tmpy=Data_GetProperty(Object,"abs_y")
	tmpr=Data_GetProperty(Object,"height")
	tmpz=Data_GetProperty(Object,"width")
	Err_CaptureObjText Gl_err,tmpvalue,tmpx,tmpy,tmpx+tmpz,tmpy+tmpr
	'MsgBox "Error"
End Function

Function Terminate()
	CloseExcel
	Log_exit
End Function

'test1 "hello"

'TestCase "D:\testing\testing\test framework\me\TestCase.xls","D:\QuicktestOR\Etc\Objectdef..."
		
		
