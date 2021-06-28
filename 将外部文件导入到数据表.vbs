
'************************************************************************************************************************
'描述：
'
'该示例将数据从外部文件导入到数据表，并通过数据表
'设置测试的参数。显示了如何通过在运行测试前设置数据表值
'来向测试发送参数。
'
'假定：
'QuickTest 中当前未打开没有保存的测试。
'有关详细信息，请参阅 Test.SaveAs 方法的示例。
'QuickTest 与测试所必需的加载项一起打开。
'有关详细信息，请参阅 Test.GetAssociatedAddins 方法的示例。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim qtOptions 'As QuickTest.RunResultsOptions ' 声明 Run Results Options 对象变量
Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见

' 打开测试
qtApp.Open "C:\Tests\Test1", False ' 打开名为“Test1”的测试

' 将数据导入到设计时数据表，然后添加新的数据
qtApp.Test.DataTable.Import "C:\Data.xls" ' 从外部文件导入数据
qtApp.Test.DataTable.ImportSheet "C:\Tables.xls", 1, "Action1" ' 导入单个工作表
qtApp.Test.DataTable.GlobalSheet("Started") = Now ' 设置测试运行的启动时间
qtApp.Test.DataTable.GlobalSheet("ParamCount") = 45 ' 使用数据表设置测试的参数

' 运行测试
Set qtOptions = CreateObject("QuickTest.RunResultsOptions") ' 创建 Results Option 对象
qtOptions.ResultsLocation = "<TempLocation>" ' 将结果的位置设置为临时位置
qtApp.Test.Run qtOptions, True ' 在继续运行自动脚本前运行测试并等待完成

' 设置设计时数据表中的其他值
qtApp.Test.DataTable.GlobalSheet("Ended") = Now ' 设置测试运行的结束时间
qtApp.Test.DataTable.GlobalSheet("RanBy") = "James" ' 设置“RanBy”单元格
qtApp.Test.Save

' 保存运行时数据表
qtApp.Test.LastRunResults.DataTable.Export "C:\Runtime.xls" ' 将运行时数据表保存到文件中

qtApp.Quit ' 退出 QuickTest
Set qtOptions = Nothing ' 释放 Run Results Options 对象
Set qtApp = Nothing ' 释放 Application 对象
