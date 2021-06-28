
'************************************************************************************************************************
'描述：
'
'该示例打开 QuickTest，而不加载任何加载项
'（仅标准 Windows 支持），并指定打开用于测试的
'应用程序。
'
'假定：
'QuickTest 中当前未打开没有保存的测试。
'有关详细信息，请参阅 Test.SaveAs 方法的示例。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim qtStdLauncher 'As QuickTest.StdLauncher ' 声明 Windows 应用程序启动程序变量
Dim qtStdApp 'As QuickTest.StdApplication ' 声明作为 StdApplication 对象变量
Dim strAdded ' 为添加的应用程序声明字符串变量

Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象

' 准备应用程序和测试
qtApp.SetActiveAddins Array() ' 删除集合中的所有加载项，以便 QuickTest 在打开时不存在任何加载项
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见
qtApp.Test.SetAssociatedAddins Array() ' 从与测试关联的加载项列表中删除所有加载项。
Set qtStdLauncher = qtApp.Test.Settings.Launchers.Item("Windows Applications") ' 返回 Windows 应用程序启动程序

qtStdLauncher.Active = True ' 指示 QuickTest 在录制会话开始时打开应用程序

' 在测试中设置应用程序
qtStdLauncher.Applications.Add "C:\Viewer.exe", "C:\" ' 添加应用程序
qtStdLauncher.Applications.Add "D:\Apps\Editor.exe", "D:\Apps" ' 添加另一个应用程序

' 保存更改并清除
qtApp.Test.SaveAs "C:\Tests\NewTest" ' 保存测试
qtApp.Quit ' 退出 QuickTest
Set qtStdLauncher = Nothing ' 释放 Windows 应用程序启动程序对象
Set qtApp = Nothing ' 释放 Application 对象
