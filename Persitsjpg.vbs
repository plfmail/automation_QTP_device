'在图片中添加文字
'输入参数filename-->文件名，strng-->在图片中添加的文字
Function JPG_TypeString(filename,strng)
	Set Jpeg=CreateObject("Persits.Jpeg")
	Jpeg.Open filename
	Jpeg.Canvas.Font.Color=&HFF0000  '红颜色
	Jpeg.Canvas.Font.Family="楷体_GB2312"
	Jpeg.Canvas.Font.Bold=True '是否加粗
	Jpeg.Canvas.Print 100,Jpeg.OriginalHeight/2,strng
	Jpeg.Save filename
	Jpeg.Close
	Set Jpeg=Nothing
End Function

'在图片中画一个椭圆
'输入参数filename-->文件名,leftlen,toplen-->左边的x,y,rightlen,buttonlen-->右边的x,y
Function JPG_DrawEllipse(filename,leftlen,toplen,rightlen,buttonlen)
	Set Jpeg=CreateObject("Persits.Jpeg")
	Jpeg.Open filename
	Jpeg.Canvas.Pen.Color=&HFF0000  '红颜色
	Jpeg.Canvas.Pen.Width=2
	Jpeg.Canvas.Brush.Solid=False '是否加粗
	Jpeg.Canvas.Ellipse leftlen,toplen,rightlen,buttonlen
	Jpeg.Save filename
	Jpeg.Close
	Set Jpeg=Nothing
End Function	

'在图片中需要标识的对象上画一个椭圆，然后在椭圆的一侧画一条线，然后在线的一侧标注信息
'输入参数filename-->文件名,leftlen,toplen-->左边的x,y,rightlen,buttonlen-->右边的x,y
Function JPG_DrawEllipseAndString(filename,leftlen,toplen,rightlen,buttonlen)
	Set Jpeg=CreateObject("Persits.Jpeg")
	Jpeg.Open filename
	Jpeg.Canvas.Pen.Color=&HFF0000  '红颜色
	Jpeg.Canvas.Pen.Width=2
	Jpeg.Canvas.Brush.Solid=False '是否加粗
	Jpeg.Canvas.Ellipse leftlen,toplen,rightlen,buttonlen '画椭圆
	'MsgBox Jpeg.OriginalHeight
	'MsgBox Jpeg.OriginalWidth
	If leftlen>Jpeg.OriginalWidth/2 Then
		tmpleft=leftlen
		tmptop=toplen+(buttonlen-toplen)/2
		If Leftlen>100 Then
			tmpright=leftlen-100
		Else
			tmpright=leftlen/2
		End If
		If toplen+(buttonlen-toplen)/2>100 Then
			tmpbutton=toplen+(buttonlen-toplen)/2-100
		Else 
			tmpbutton=toplen+(buttonlen-toplen)/2+100
		End If
	Else
		tmpleft=rightlen
		tmptop=toplen+(buttonlen-toplen)/2
		If rightlen+100+Len(strng)*2>Jpeg.OriginalWidth Then
			tmpright=Jpeg.OriginalWidth-Len(strng)*2
			'未完成
	
	
	
	
	Jpeg.Save filename
	Jpeg.Close
	Set Jpeg=Nothing
End Function	 
	
