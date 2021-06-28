
'************************************************************************************************************************
'描述：
'
'该示例打开 QuickTest，并配置其 Windows 应用程序选项。
'
'假定：
'未打开 QuickTest。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim qtOptions 'As QuickTest.WindowsAppsOptions ' 声明 Windows Applications Options 对象变量

Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见

Set qtOptions = qtApp.Options.WindowsApps ' 返回 Windows Applications Options 对象

' 配置 Windows 应用程序选项
qtOptions.AttachedTextArea = "BottomLeft" ' 设置搜索附加文本的点
qtOptions.AttachedTextRadius = 50 ' 设置搜索附加文本的最大距离
qtOptions.ExpandMenuToRetrieveProperties = True ' 设置运行期间检索属性前打开菜单对象
qtOptions.NonUniqueListItemRecordMode = "ByIndex" ' 如果存在多个列表或树项名称相同，则设置录制各自的索引
qtOptions.RecordOwnerDrawnButtonAs = "CheckBoxes" ' 设置标识或录制自定义按钮作为复选框
qtOptions.ForceEnumChildWindows = True ' 录制和运行测试时枚举所有子窗口。
qtOptions.ClickEditBeforeSetText = True ' 在编辑框中插入文本前，执行单击操作在编辑框中设置焦点。
qtOptions.VerifyMenuInitEvent = False ' 在菜单控件上录制操作前忽略菜单初始化事件。


Set qtOptions = Nothing ' 释放 Windows Applications Options 对象
Set qtApp = Nothing ' 释放 Application 对象

