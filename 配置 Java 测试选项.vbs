'************************************************************************************************************************
'描述：
'
'该示例打开 QuickTest，并配置与
'“选项”对话框和“高级 Java 选项”对话框的“Java”选项卡相对应的选项。
'
'假定：
'QuickTest 中当前未打开没有保存的测试。
'有关详细信息，请参阅 Test.SaveAs 方法的示例。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim qtJavaOptions 'As QuickTest.JavaOptions ' 声明 Java Options 对象变量

Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.SetActiveAddins Array("Java") ' 加载 Java 加载项
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见
Set qtJavaOptions = qtApp.Options.Java ' 返回 Java Options 对象

' 配置 Java 选项
qtJavaOptions.RecordListByIndex = False '按名称录制列表项
qtJavaOptions.RecordComboByIndex = True '按索引录制组合框选项
qtJavaOptions.RecordTreeByIndex = False '按名称录制树节点选项
qtJavaOptions.RecordTabByIndex = False '按名称录制选项卡式窗格控件中的选项卡选项
qtJavaOptions.AttachedTextRadius = "100" '在 100 像素的半径范围内搜索组件的附加文本
qtJavaOptions.DeviceReplay = "DragDrop" '为拖放事件使用设备级回放
qtJavaOptions.AWTEventModel = "Auto" ' 启用 QuickTest 自动选择模型，用于向对象发送事件
qtJavaOptions.AnalogTableRecording = True '以模拟模式录制表操作

Set qtJavaOptions = Nothing ' 释放 Java Options 对象
Set qtApp = Nothing ' 释放 Application 对象
