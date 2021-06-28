
'************************************************************************************************************************
'描述：
'
'该示例打开新的测试，并为该测试配置资源。
'
'假定：
'QuickTest 中当前未打开没有保存的测试。
'有关详细信息，请参阅 Test.SaveAs 方法的示例。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim qtTestResources 'As QuickTest.Resources ' 声明 Resources 对象变量

Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见
qtApp.New ' 打开新的测试

' 返回 Resources 对象
Set qtTestResources = qtApp.Test.Settings.Resources

' 指定外部数据表文件和共享的对象库
qtTestResources.DataTablePath = "C:\Resources\Default.xls"
qtTestResources.ObjectRepositoryPath = "C:\Resources\Resource.mtr"
' 使该共享库为所有新的测试的默认库
qtTestResources.SetObjectRepositoryAsDefault

Set qtTestResources = Nothing ' 释放 Resources 对象
Set qtApp = Nothing ' 释放 Application 对象
