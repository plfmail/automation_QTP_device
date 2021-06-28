
'************************************************************************************************************************
'描述：
'
'该示例配置常规 QuickTest 选项。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量

Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动应用程序
qtApp.Visible = True ' 使 QuickTest 可见

' 配置选项
qtApp.Options.AutoGenerateWith = True
qtApp.Options.WithGenerationLevel = 3
qtApp.Options.DisableVORecognition = True
qtApp.Options.SaveLoadAndMonitorData = False
qtApp.Options.TimeToActivateWinAfterPoint = 600
qtApp.Options.RestoreLayout

Set qtApp = Nothing ' 释放 Application 对象
