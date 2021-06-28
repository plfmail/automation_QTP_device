'************************************************************************************************************************
'描述：
'
'该示例打开 QuickTest，使用 Folders 集合配置
'用来解析相对路径的搜索路径。
'
'假定：
'QuickTest 中当前未打开没有保存的测试。
'有关详细信息，请参阅 Test.SaveAs 方法的示例。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim strPath
Dim lngPosition

' 打开 QuickTest
Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见

' 打开测试
qtApp.Open "C:\Tests\Test1", True, False ' 以只读模式打开测试

' 找到“Folder1”，如果它不在集合中，则添加它。
strPath = qtApp.Folders.Locate("..\..\Folders\Folder1") ' 找到“Folder1”文件夹
If qtApp.Folders.Find(strPath) = -1 Then ' 如果未在集合中找到文件夹
    qtApp.Folders.Add strPath, 1 ' 将文件夹添加到集合
End If

' 如果推入了当前的文件夹 - 将其还原到位置 1
If qtApp.Folders.Item(2) = "<Current Test>" Then ' 如果当前测试文件夹位于位置 2
    qtApp.Folders.MoveToPos 1, 2 ' 调换前两个文件夹的顺序
End If

' 如果“Folder2”在集合中 - 将其删除。
lngPosition = CLng(qtApp.Folders.Find("C:\Folders\Folder2")) ' 搜索“Folder2”文件夹
If intPosition <> -1 Then ' 如果在集合中找到文件夹
    qtApp.Folders.Remove lngPosition ' 将其删除
End If

Set qtApp = Nothing ' 释放 Application 对象

