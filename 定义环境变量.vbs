
'************************************************************************************************************************
'描述：
'
'该示例通过在运行测试前定义环境变量来向测试发送参数。
'测试描述中还列出了每次运行测试的用户和操作系统。
'
'假定：
'QuickTest 中当前未打开没有保存的测试。
'有关详细信息，请参阅 Test.SaveAs 方法的示例。
'打开 QuickTest 时，将加载测试所必需的加载项。
'有关详细信息，请参阅 Test.GetAssociatedAddins 方法的示例。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim qtOptions 'As QuickTest.RunResultsOptions ' 声明 Run Results Options 对象变量
Dim strOS
Dim strUserName

Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见

' 打开测试
qtApp.Open "C:\Tests\Test1", False ' 打开名为“Test1”的测试

' 设置一些环境变量
qtApp.Test.Environment.Value("Root") = "C:\" ' 设置“Root”环境变量。注意：'“Value”是默认属性，且可以省略。
qtApp.Test.Environment.Value("Password") = "not4you" ' 设置“Password”环境变量
qtApp.Test.Environment.Value("Days") = 14 ' 设置“Days”环境变量

' 运行测试
Set qtOptions = CreateObject("QuickTest.RunResultsOptions") ' 创建 Results Option 对象
qtOptions.ResultsLocation = "<TempLocation>" ' 将结果的位置设置为临时位置
qtApp.Test.Run qtOptions, True ' 在继续运行自动脚本前运行测试并等待完成

' 将在环境中存储的数据连接到描述
strOS = qtApp.Test.Environment.Value("OS") ' 返回“OS”环境变量
strUserName = qtApp.Test.Environment.Value("UserName") ' 返回“UserName”环境变量
qtApp.Test.Description = qtApp.Test.Description & vbNewLine & _
       "Tested on: " & Now & vbNewLine & _
       "Platform: " & strOS & vbNewLine & _
       "By user: " & strUserName & vbNewLine & _
       "-------------------------" & vbNewLine ' 向测试的描述添加当前时间、操作系统和用户名

qtApp.Test.Save ' 保存测试

qtApp.Quit ' 退出 QuickTest
Set qtOptions = Nothing ' 释放 Run Results Options 对象
Set qtApp = Nothing ' 释放 Application 对象
