
'************************************************************************************************************************
'描述：
'
'该示例打开测试，配置运行选项和设置，
'运行该测试，然后检查测试运行的结果。
'
'假定：
'QuickTest 中当前未打开没有保存的测试。
'有关详细信息，请参阅 Test.SaveAs 方法的示例。
'打开 QuickTest 时，将加载测试所必需的加载项。
'有关详细信息，请参阅 Test.GetAssociatedAddins 方法的示例。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim qtTest 'As QuickTest.Test ' 声明 Test 对象变量
Dim qtResultsOpt 'As QuickTest.RunResultsOptions ' 声明 Run Results Options 对象变量

Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见

' 设置 QuickTest 运行选项
qtApp.Options.Run.CaptureForTestResults = "OnError"
qtApp.Options.Run.RunMode = "Fast"
qtApp.Options.Run.ViewResults = False

qtApp.Open "C:\Tests\Test1", True ' 以只读模式打开测试

' 为测试设置运行设置
Set qtTest = qtApp.Test
qtTest.Settings.Run.IterationMode = "rngIterations" ' 仅运行循环 2 到 4
qtTest.Settings.Run.StartIteration = 2
qtTest.Settings.Run.EndIteration = 4
qtTest.Settings.Run.OnError = "NextStep" ' 指示 QuickTest 在发生错误时执行下一步骤

Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions") ' 创建 Run Results Options 对象
qtResultsOpt.ResultsLocation = "C:\Tests\Test1\Res1" ' 设置结果位置

qtTest.Run qtResultsOpt ' 运行测试

MsgBox qtTest.LastRunResults.Status ' 检查测试运行的结果
qtTest.Close ' 关闭测试

Set qtResultsOpt = Nothing ' 释放 Run Results Options 对象
Set qtTest = Nothing ' 释放 Test 对象
Set qtApp = Nothing ' 释放 Application 对象
