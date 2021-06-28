
'************************************************************************************************************************
'描述：
'
'该示例启动 QuickTest，打开新的测试，并将其配置用于在 SAP GUI for Windows 应用程序上
'录制和运行
'
'假定：
'未打开 QuickTest。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象

qtApp.SetActiveAddins Array("SAP") ' 加载 SAP 加载项。
qtApp.Launch ' 启动 QuickTest
qtApp.New ' 打开新的测试

' 配置“SAP 录制和运行”设置，以便打开具有这些设置的 SAP GUI for Windows 应用程序：
qtApp.Test.Settings.Launchers("SAP").Active = True 'Invoke a new SAP Gui for Windows session when recording begins
qtApp.Test.Settings.Launchers("SAP").Server = "R/3 Enterprise" '启动 SAP GUI for Windows，并连接到“R/3 Enterprise”服务器
qtApp.Test.Settings.Launchers("SAP").AutoLogon = True '使用以下登录详细信息执行自动登录
qtApp.Test.Settings.Launchers("SAP").Client = "800" 'SAP 客户端的编号
qtApp.Test.Settings.Launchers("SAP").User = "QA01" 'SAP 服务器的用户名
qtApp.Test.Settings.Launchers("SAP").Password = "3f5aea819b0239" '设置密码为一个加密字符串
qtApp.Test.Settings.Launchers("SAP").Language = "EN" '用户界面使用英语
qtApp.Test.Settings.Launchers("SAP").RememberPassword = True 'Save logon password for use in future test runs
qtApp.Test.Settings.Launchers("SAP").CloseOnExit = True '退出该测试时关闭该 SAP GUI for Windows 会话
qtApp.Test.Settings.Launchers("SAP").IgnoreExistingSessions = True '不要在录制或运行会话开始之前已经打开的任何 SAP 会话上录制或运行测试

qtApp.Visible = True ' 使 QuickTest 应用程序可见
Set qtApp = Nothing ' 释放 Application 对象
