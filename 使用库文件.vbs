'************************************************************************************************************************
'描述：
'
'该示例打开测试，配置测试的库的集合
'并保存测试。
'
'假定：
'QuickTest 中当前未打开没有保存的测试。
'有关详细信息，请参阅 Test.SaveAs 方法的示例。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim qtLibraries 'As QuickTest.TestLibraries ' 声明测试的库集合变量
Dim lngPosition

' 打开 QuickTest
Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 设置 QuickTest 可见

' 打开测试并获取其库集合
qtApp.Open "C:\Tests\Test1", False, False ' 打开测试
Set qtLibraries = qtApp.Test.Settings.Resources.Libraries ' 获取库集合对象

' 如果 Utilities.vbs 不在集合中，则添加它
If qtLibraries.Find("C:\Utilities.vbs") = -1 Then ' 如果集合中找不到库
    qtLibraries.Add "C:\Utilities.vbs", 1 ' 向集合添加库
End If

' 如果推入了 Math.vbs - 将其还原到位置 1
If qtLibraries.Count > 1 And qtLibraries.Item(2) = "C:\Math.vbs" Then ' 如果存在多个库且 Math.vbs 位于位置 2
    qtLibraries.MoveToPos 1, 2 ' 在前两个库之间进行切换
End If

' 如果 Debug.vbs 不在集合中 - 将其删除
lngPosition = qtLibraries.Find("C:\Debug.vbs") ' 尝试查找 Debug.vbs 库
If lngPosition <> -1 Then ' 如果在集合中找到库
    qtLibraries.Remove lngPosition ' 将其删除
End If

' 将新库的配置设置为默认配置
qtLibraries.SetAsDefault ' 将与测试关联的库文件设置为新库的默认库文件

'保存测试并关闭 QuickTest
qtApp.Test.Save ' 保存测试
qtApp.Quit ' 退出 QuickTest

Set qtLibraries = Nothing ' 释放测试的库集合
Set qtApp = Nothing ' 释放 Application 对象
