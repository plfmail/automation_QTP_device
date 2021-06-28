'************************************************************************************************************************
'描述：
'
'该示例打开一个测试，并加载所有与测试关联的加载项。
'
'假定：
'QuickTest 中当前未打开没有保存的测试。
'有关详细信息，请参阅 Test.SaveAs 方法的示例。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim blnNeedChangeAddins ' 声明一个标志，用于指明当前是否已加载与测试关联的加载项
Dim arrTestAddins ' 声明变量，用于存储与测试关联的加载项

Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象

arrTestAddins = qtApp.GetAssociatedAddinsForTest("C:\Tests\Test1") ' 创建一个数组，用于包含与该测试关联的加载项的列表

' 检查是否已加载所有必需的加载项
blnNeedChangeAddins = False ' 假定无需作任何更改
For Each testAddin In arrTestAddins ' 循环与测试关联的加载项列表
    If qtApp.Addins(testAddin).Status <> "Active" Then ' 如果存在未加载的关联加载项
        blnNeedChangeAddins = True ' 指明需要对加载的加载项进行更改
        Exit For ' 退出循环
    End If
Next

If qtApp.Launched And blnNeedChangeAddins Then
        qtApp.Quit ' 如果有必要进行更改，则退出 QuickTest，修改已加载的加载项
End If

If blnNeedChangeAddins Then
    Dim blnActivateOK
    blnActivateOK = qtApp.SetActiveAddins(arrTestAddins, errorDescription) ' 加载与测试关联的加载项并检查它们是否已加载成功。
    If Not blnActivateOK Then ' 如果在加载加载项时发生问题
        MsgBox errorDescription ' 显示包含错误的消息
    WScript.Quit ' 并结束自动程序。
    End If
End If

If Not qtApp.Launched Then ' 如果尚未打开 QuickTest
    qtApp.Launch ' 启动 QuickTest（已加载正确的加载项）
End If
qtApp.Visible = True ' 使 QuickTest 应用程序可见

qtApp.Open "C:\Tests\Test1" ' 打开测试
Set qtApp = Nothing ' 释放 Application 对象

