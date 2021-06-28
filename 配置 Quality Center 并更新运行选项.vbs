
'************************************************************************************************************************
'描述：
'
'该示例连接到 Quality Center 项目，打开测试（如果可用，将其签出），
'更新 Active Screen 值并测试对象描述，以及如果可用，
'将已修改的测试签入到 Quality Center 项目。
'
'假定：
'test1 测试尚未签出。
'QuickTest 中当前未打开没有保存的测试。
'有关详细信息，请参阅 Test.SaveAs 方法的示例。
'打开 QuickTest 时，将加载测试所必需的加载项。
'有关详细信息，请参阅 Test.GetAssociatedAddins 方法的示例。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim qtUpdateRunOptions 'As QuickTest.UpdateRunOptions ' 声明 Update Run Options 对象变量
Dim qtRunResultsOptions 'As QuickTest.RunResultsOptions ' 声明 Run Results Options 对象变量
Dim blsSupportsVerCtrl ' 声明一个标志，用于指明版本控制支持

Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见

' 在具有版本控制的 Quality Center 上更改测试
qtApp.TDConnection.Connect "http://tdserver/tdbin", _
              "MY_DOMAIN", "My_Project", "James", "not4you", False ' 连接到 Quality Center

If qtApp.TDConnection.IsConnected Then ' 如果连接成功
    blsSupportsVerCtrl = qtApp.TDConnection.SupportVersionControl ' 检查项目是否支持版本控制

    qtApp.Open "[QualityCenter] Subject\tests\test1", False ' 打开测试
    If blsSupportsVerCtrl Then ' 如果项目支持版本控制
        qtApp.Test.CheckOut ' 签出测试
    End If

    ' 准备 UpdateRunOptions 对象
    Set qtUpdateRunOptions = CreateObject("QuickTest.UpdateRunOptions") ' 创建 Update Run Options 对象
    ' 设置“更新运行”选项：更新 Active Screen 和测试对象描述。不要更新检查点值
    qtUpdateRunOptions.UpdateActiveScreen = True
    qtUpdateRunOptions.UpdateCheckpoints = False
    qtUpdateRunOptions.UpdateTestObjectDescriptions = True

    ' 准备 RunResultsOptions 对象
    Set qtRunResultsOptions = CreateObject("QuickTest.RunResultsOptions") ' 创建 Run Results Options 对象
    qtRunResultsOptions.ResultsLocation = "<TempLocation>" ' 设置临时结果位置

    '更新测试
    qtApp.Test.UpdateRun qtUpdateRunOptions, qtRunResultsOptions ' 以“更新运行”模式运行测试
    qtApp.Test.Description = qtApp.Test.Description & vbNewLine & _
                              "Updated: " & Now ' 记录测试的描述中的更新（“测试设置”>“属性”选项卡）

    qtApp.Test.Save ' 保存测试

    If blsSupportsVerCtrl And qtApp.Test.VerCtrlStatus = "CheckedOut" Then ' 如果已签出测试
        qtApp.Test.CheckIn ' 将其签入
    End If

    qtApp.TDConnection.Disconnect ' 断开与 Quality Center 的连接
Else
    MsgBox "Cannot connect to Quality Center" ' 如果连接不成功，则显示一则错误消息。
End If

qtApp.Quit ' 退出 QuickTest
Set qtUpdateRunOptions = Nothing ' 释放 Update Run Options 对象
Set qtRunResultsOptions = Nothing ' 释放 Run Results Options 对象
Set qtApp = Nothing ' 释放 Application 对象

