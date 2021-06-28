
'************************************************************************************************************************
'描述：
'
'该示例启动 QuickTest，打开新的测试，并将其配置用于在 Web 应用程序上录制和运行。
'
'假定：
'未打开 QuickTest。
'************************************************************************************************************************

Dim qtApp 'As QuickTest.Application ' 声明 Application 对象变量
Set qtApp = CreateObject("QuickTest.Application") ' 创建应用程序对象

qtApp.SetActiveAddins Array("Web") ' 激活 Web 加载项
qtApp.Launch ' 启动 QuickTest
qtApp.New ' 打开新的测试

' 配置 Web 应用程序使用该测试
qtApp.Test.Settings.Launchers("Web").Active = True
qtApp.Test.Settings.Launchers("Web").Browser = "IE"
qtApp.Test.Settings.Launchers("Web").Address = "http://newtours.mercuryinteractive.com "
qtApp.Test.Settings.Launchers("Web").CloseOnExit = True

' 配置 Active Screen 访问设置
qtApp.Test.Settings.Web.ActiveScreenAccess.UserName = "user1"
qtApp.Test.Settings.Web.ActiveScreenAccess.Password = "mypassword"

' 配置其他 Web 设置
qtApp.Test.Settings.Web.BrowserNavigationTimeout = 60000
qtApp.Test.Settings.Web.NextPageIfObjNotFound = True

qtApp.Visible = True ' 使 QuickTest 应用程序可见
Set qtApp = Nothing ' 释放 Application 对象
