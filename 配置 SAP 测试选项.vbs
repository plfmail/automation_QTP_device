
'************************************************************************************************************************
'描述：
'
'该示例打开 QuickTest，并配置用于在 SAP GUI for Windows 应用程序上
'录制和运行测试的选项。
'
'假定：
'QuickTest 中当前未打开没有保存的测试。
'有关详细信息，请参阅 Test.SaveAs 方法的示例。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量

Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.SetActiveAddins Array("SAP") ' 加载 SAP 加载项
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见

' 配置高级 SAP 选项
qtApp.Options.SAP.AutoParameterizeTables = True '指定在编辑 SAP 表和网格控件中的数据时应录制单个 Input 语句。这将创建一个数据表，其中包含在单次服务器通信过程中设置的所有单元格的数据。
qtApp.Options.SAP.RecordStatusBar = True '指明每次 SAP 状态栏显示消息时 QuickTest 是否录制一个同步步骤并捕获对应的 Active Screen 内容。
qtApp.Options.SAP.SessionCleanup = True '指明测试或组件关闭时是否关闭由它们打开的所有 SAP 会话。
qtApp.Options.SAP.RecordResetMethod = False '指明在使用 SAP 的“自动登录”选项打开的每个录制会话开始时，QuickTest 是否录制一个 Reset 方法。
qtApp.Options.SAP.ShowLowSpeedWarnings = True '指明每次 SAP GUI 服务器设置为使用“低速连接”选项时 QuickTest 是否显示警告消息。
qtApp.Options.SAP.ShowDisabledConnectionWarnings = True '指明每次禁用“SAP GUI 脚本接口”时 QuickTest 是否显示警告消息。
qtApp.Options.SAP.UseSapGuiScriptingForHTML = True '指明 QuickTest 是否使用“SAP GUI 脚本接口”来录制 HTML 元素。

Set qtApp = Nothing ' 释放 Application 对象
