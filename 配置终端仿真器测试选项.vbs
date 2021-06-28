
'************************************************************************************************************************
'描述：
'
'该示例打开 QuickTest 并配置“终端仿真器”选项。
'
'假定：
'未打开 QuickTest。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim qtTeOptions 'As QuickTest.TeOptions ' 声明 TE Options 对象变量

Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.SetActiveAddins Array("Terminal Emulators") ' 激活终端仿真器加载项
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见
Set qtTeOptions = qtApp.Options.TE ' 返回 TE Options 对象

' 配置“终端仿真器”选项

' 设置全局“终端仿真器”选项（适用于所有仿真器）

qtTeOptions.ScreenTitleRow = 1 ' 设置屏幕标题的位置
qtTeOptions.ScreenTitleCol = 1
qtTeOptions.ScreenTitleLength = 30

qtTeOptions.AutoAdvance = False ' 指定仿真器不支持自动前进字段
qtTeOptions.AutoSyncKeys = "13" ' 录制测试时，录制每次按下 Enter 键时的同步步骤
qtTeOptions.RecordMenusAndPopups = True '启用录制仿真器弹出信息和菜单
qtTeOptions.RecordCursorPosition = True '录制光标位置的步骤
qtTeOptions.UsePropertyPattern = True ' 使用默认属性模式文件
qtTeOptions.PropertyPatternsFile = "C:\Program Files\Mercury Interactive\QuickTest Professional\Dat\PropertyPatternConfigTE.xml"

' 指定当前仿真器
qtTeOptions.CurrentEmulator = "Host On-Demand 8.0" ' 设置当前仿真器为 Host On-Demand 8.0

' 设置特定于当前仿真器的选项
qtTeOptions.Protocol = "autodetect" '使 QuickTest 检测会话协议
qtTeOptions.BlankLines = 0 ' 仿真器窗口底部没有空白行。
qtTeOptions.CodePage = 0 ' 使用默认代码页转换
qtTeOptions.HllapiDllName = "C:\Program Files\IBM\EHLLAPI\pcshll32.dll" ' 指定要使用的 HLLAPI dll
qtTeOptions.HllapiProcName = "hllapi" ' 指定要使用的 HLLAPI 函数
qtTeOptions.VerifyHllapiDllPath = True ' 如果找不到 HLLAPI dll，则显示警告消息
qtTeOptions.ScreenLabelUseAllChars = True ' 使用受保护和不受保护的字段标识屏幕标签。
qtTeOptions.WindowTitlePrefix = "MyTerminal" ' 根据终端窗口的窗口标题前缀来标识终端窗口。
qtTeOptions.TrailingMode = True ' 当以上下文有关的模式进行录制时，剪裁白色字符
qtTeOptions.TrailingFieldLength = 5 ' 如果字段以五个黑色字符开头。
qtTeOptions.BeepOnSync = False ' 不要在每一同步步骤后发出声音。
qtTeOptions.UseKeyEvent "@R", "17;52" ' 使用 CTRL+R 键盘事件回放 TE_RESET 键。
qtTeOptions.SyncTime = 200 ' 等待 200 毫秒再检查仿真器状态

' 清除
Set qtTeOptions = Nothing ' 释放 TE Options 对象 ' 释放 Te Options 对象
Set qtApp = Nothing ' 释放 Application 对象 ' 释放 Application 对象
