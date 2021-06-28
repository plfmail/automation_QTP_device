
'************************************************************************************************************************
'描述：
'
'该示例启动 QuickTest，打开新的测试，并将其配置用于在具有自动登录功能活动的 Siebel 7.7 应用程序上进行录制和运行。
'
'假定：
'未打开 QuickTest。
'************************************************************************************************************************
Dim qtApp 'As Application ' 声明 Application 对象变量
Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.SetActiveAddins Array("Siebel") ' 激活 Siebel 加载项
qtApp.Launch ' 启动 QuickTest
qtApp.New ' 打开新的测试
' 配置 Siebel 应用程序版本来使用该测试
qtApp.Test.Settings.Launchers("Siebel").Active = True
qtApp.Test.Settings.Launchers("Siebel").Version = "77"
' 指定 URL 并配置浏览器设置来使用该测试
qtApp.Test.Settings.Launchers("Siebel").Address = "http://Siebel_application.url"
qtApp.Test.Settings.Launchers("Siebel").LogoutOnExit = False
qtApp.Test.Settings.Launchers("Siebel").CloseOnExit = True
' 配置自动登录参数来使用该测试
qtApp.Test.Settings.Launchers("Siebel").AutoLogin = True
qtApp.Test.Settings.Launchers("Siebel").User = "username"
qtApp.Test.Settings.Launchers("Siebel").Password = "mypassword"
' 配置 Siebel 7.7 应用程序特定的高级超时和访问设置
qtApp.Test.Settings.Launchers("Siebel").SiebAutomationRequestTimeout = 120
qtApp.Test.Settings.Launchers("Siebel").SiebAutomationAccessCode = "aCode"
qtApp.Visible = True ' 使 QuickTest 应用程序可见
Set qtApp = Nothing ' 释放 Application 对象
