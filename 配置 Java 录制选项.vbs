
'************************************************************************************************************************
'描述：
'
'该示例启动 QuickTest，打开新的测试，并配置
'其用于在 Java 应用程序上录制和运行。
'
'假定：
'未打开 QuickTest。
'************************************************************************************************************************

Dim qtApp 'As QuickTest.Application ' 声明 Application 对象变量
Set qtApp = CreateObject("QuickTest.Application") ' 创建应用程序对象

qtApp.SetActiveAddins Array("Java") ' 激活 Java 加载项
qtApp.Launch ' 启动 QuickTest
qtApp.New ' 打开新的测试

' 配置 Java 应用程序使用该测试
qtApp.Test.Settings.Launchers("Java").Active = True
qtApp.Test.Settings.Launchers("Java").CommandLine = "C:\j2sdk1.4.2\bin\java.exe -jar C:\j2sdk1.4.2\demo\jfc\SwingSet2\SwingSet2.jar"
qtApp.Test.Settings.Launchers("Java").WorkingDirectory = "C:\j2sdk1.4.2\demo\jfc\SwingSet2"

qtApp.Visible = True ' 使 QuickTest 应用程序可见
Set qtApp = Nothing ' 释放 Application 对象
