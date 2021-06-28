'************************************************************************************************************************
'描述：
'
'该示例打开 QuickTest，并配置其 Web 选项。
'
'假定：
'QuickTest 中当前未打开没有保存的测试。
'有关详细信息，请参阅 Test.SaveAs 方法的示例。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim qtWebOptions 'As QuickTest.WebOptions ' 声明 Web Options 对象变量

Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.SetActiveAddins Array("Web") ' 激活 Web 加载项
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见
Set qtWebOptions = qtApp.Options.Web ' 返回 Web Options 对象

' 配置 Web 选项
qtWebOptions.AddToPageLoadTime = 30 ' 将添加到页面的加载时间设置为 30 秒
qtWebOptions.CheckBrokenLinks = True ' 设置仅检查当前中断链接的主机

' 配置高级 Web 选项
qtWebOptions.EnableBrowserResize = False ' 设置浏览器在打开时为默认大小
qtWebOptions.RunUsingSourceIndex = True ' 设置使用源索引属性（以获得更佳的性能）
qtWebOptions.RunOnlyClick = True ' 设置运行单击事件，如 MouseDown、MouseUp 和 Click
qtWebOptions.BrowserCleanup = True ' 设置在完成测试或循环时关闭所有打开的浏览器
qtWebOptions.RecordByWinMouseEvents = "OnClick OnMouseDown" ' 指明哪些事件使用标准 Windows 事件
qtWebOptions.RecordAllNavigations = True ' 设置在每次 URL 更改时录制导航
qtWebOptions.RecordMouseDownAndUpAsClick = False ' 设置录制 MouseDown 和 MouseUp，而不录制 Click
qtWebOptions.RecordCoordinates = False ' 指示 QuickTest 不录制实际坐标
If qtWebOptions.PageCreationMode = "URL" Then ' 如果当前选定了优化 Page 对象创建（URL 模式），则
    qtWebOptions.CreatePageUsingNonUserData = "Get" ' 指示如果使用 Get 转换方法，则 QuickTest 忽略非用户数据
    qtWebOptions.CreatePageUsingUserData = "Get Post" ' 指示如果使用 Get/Post 转换方法，则 QuickTest 忽略用户数据
    qtWebOptions.CreatePageUsingAdditionalInfo = False ' 指示 QuickTest 不使用其他属性来标识现有 Page
End If
If qtWebOptions.FrameCreationMode = "URL" Then ' 如果当前选定了 Frame 对象创建（URL 模式），则
    qtWebOptions.CreateFrameUsingNonUserData = "Get" ' 指示如果使用 Get 转换方法，则 QuickTest 忽略非用户数据
    qtWebOptions.CreateFrameUsingUserData = "Get Post" ' 指示如果使用 Get/Post 转换方法，则 QuickTest 忽略用户数据
    qtWebOptions.CreateFrameUsingAdditionalInfo = False ' 指示 QuickTest 不使用其他属性标识现有 Frame
End If

Set qtWebOptions = Nothing ' 释放 Web Options 对象
Set qtApp = Nothing ' 释放 Application 对象
