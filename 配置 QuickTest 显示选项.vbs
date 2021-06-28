'************************************************************************************************************************
'描述：
'
'该示例配置 QuickTest 视图和窗格，以便以可见模式运行 QuickTest。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动 QuickTest


qtApp.ActivateView "ExpertView" ' 显示专家视图
qtApp.ShowPaneScreen "ActiveScreen", True ' 显示 Active Screen 窗格
qtApp.ShowPaneScreen "DataTable", False ' 隐藏“数据表”窗格
qtApp.ShowPaneScreen "DebugViewer", True ' 显示“调试查看器”窗格
qtApp.WindowState = "Maximized" ' 最大化 QuickTest 窗口
qtApp.Visible = True ' 使 QuickTest 窗口可见

Set qtApp = Nothing ' 释放 Application 对象
