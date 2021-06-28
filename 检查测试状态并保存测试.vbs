
'************************************************************************************************************************
'描述：
'
'该示例保存已修改的测试，以便可以打开新的测试
'或退出应用程序，而不丢失任何未保存的信息。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量

Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动 QuickTest（如果未启动）
qtApp.Visible = True ' 使其可见

' 保存当前测试并根据需要决定是否打开一个新的测试
If qtApp.Test.Modified Then ' 如果修改了测试
    If qtApp.Test.IsNew Then ' 如果是新的测试（无标题）
        qtApp.Test.SaveAs "C:\Temp\TempTest" ' 使用临时名称保存测试（覆盖现有的临时测试）
    Else ' 如果存在测试（具有名称）
        qtApp.Test.Save ' 保存更改
    End If
End If

If Not qtApp.Test.IsNew Then ' 如果当前测试不是新的测试
    qtApp.New ' 打开新的测试
End If

Set qtApp = Nothing ' 释放 Application 对象
