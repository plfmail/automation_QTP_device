Function Err_CapturePage(strng)
	tmpFile=Err_GenFileName
	Err_JudgeFile tmpFile
	Desktop.CaptureBitmap tmpFile
	Call JPG_TypeString(tmpFile,strng)
End Function

Private Function Err_GenFileName()
	Err_GenFileName="c:\temp\"&RamdomNumber(1,10000)&Gl_CaseName&Gl_Step&".bmp"
End Function

Function Err_CaptureObjText(obj,strng,leftlen,toplen,rightlen,buttonlen)
	tmpFile=ErrGenFileName
	Desktop.CaptureBitmap tmpFile
	Call JPG_DrawEllipseAndString(tmpFile,strng,leftlen,toplen,rightlen,buttonlen)
End Function

Private Function Err_JudgeFile(filename)
	Set fileobj=CreateObject("Scripting FileSystemObject")
	If fileobj.FileExists(filename) Then
		fileobj.DeleteFile filename
	End If
	Set fileobj=Nothing
End Function

Function Err_DealErr(obj)
	On Error Goto Next
		tmpx=Data_GetProperty(obj,"abs_x")
		tmpy=Data_GetProperty(obj,"abs_y")
		tmpr=Data_GetProperty(obj,"height")
		tmpz=Data_GetProperty(obj,"width")
		tmpvalue=Gl_err
		Err_CaptureObjText object1,tmpvalue,tmpx,tmpy,tmpx+tmpz,tmpy+tmpr
	On Error Goto 0
End Function
