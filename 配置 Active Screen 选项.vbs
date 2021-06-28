
'************************************************************************************************************************
'描述：
'
'该示例打开 QuickTest，并将其 Active Screen 选项配置为最小级别。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim qtActiveScreenOpt 'As QuickTest.ActiveScreenOptions ' 声明 Active Screen Options 对象变量
Dim qtWebActiveScreen 'As QuickTest.WebActiveScreen ' 声明 Web Active Screen 对象变量

' 设置对象变量并打开 QuickTest
Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见
Set qtActiveScreenOpt = qtApp.Options.ActiveScreen ' 返回 Active Screen Options 对象

If qtActiveScreenOpt.CaptureLevel <> "None" Then ' 如果当前的捕获级别不为“None”

    ' 配置 Active Screen 常规选项
    qtActiveScreenOpt.CaptureLevel = "Minimum" ' 捕获属性仅适用于被录制对象

    ' 配置与 Web 相关的 Active Screen 选项
    Set qtWebActiveScreen = qtActiveScreenOpt.Web ' 返回 Web Active Screen 对象
    qtWebActiveScreen.ActiveScripts = "Disabled" ' 防止在 Active Screen 加载页面时运行脚本
    qtWebActiveScreen.CaptureOriginalHTMLSource = True ' 在运行任何脚本前，捕获 Web 页的原始 HTML 源
    qtWebActiveScreen.LoadActiveXControls = False ' 不允许在 Active Screen 窗格中加载 ActiveX 控件
    qtWebActiveScreen.LoadImages = False ' 不允许在 Active Screen 窗格中加载图像
    qtWebActiveScreen.LoadJavaApplets = False ' 不允许在 Active Screen 窗格中加载 Java Applet
    qtWebActiveScreen.LoadingTimeout = 20 ' 将 Active Screen 加载页面的最长时间设置为 20 秒
    Set qtWebActiveScreen = Nothing ' 释放 Web Active Screen 对象

End If

Set qtActiveScreenOpt = Nothing ' 释放 Active Screen Options 对象
Set qtApp = Nothing ' 释放 Application 对象
