'************************************************************************************************************************
'描述：
'
'该示例使用 Recovery 集合指定一组
'恢复场景来关联新的测试。
'
'假定：
'QuickTest 中当前未打开没有保存的测试。
'有关详细信息，请参阅 Test.SaveAs 方法的示例。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim qtTestRecovery 'As QuickTest.Recovery ' 声明 Recovery 对象变量
Dim intIndex ' 声明索引变量

' 打开 QuickTest 并准备对象变量
Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动 QuickTest
qtApp.New ' 打开新的测试
qtApp.Visible = True ' 使 QuickTest 应用程序可见
Set qtTestRecovery = qtApp.Test.Settings.Recovery ' 返回当前测试的 Recovery 对象

If qtTestRecovery.Count > 0 Then ' 如果为测试指定了某些默认场景
    qtTestRecovery.RemoveAll ' 删除它们
End If

' 添加恢复场景
qtTestRecovery.Add "C:\Recovery.qrs", "ErrMessage", 1 ' 将“ErrMessage”场景添加为第一个场景
qtTestRecovery.Add "C:\Recovery.qrs", "AppCrash", 2 ' 将“AppCrash”场景添加为第二个场景
qtTestRecovery.Add "C:\Recovery.qrs", "ObjDisabled", 3 ' 将“ObjDisabled”场景添加为第三个场景

' 启用所有场景
For intIndex = 1 To qtTestRecovery.Count ' 循环场景
    qtTestRecovery.Item(intIndex).Enabled = True ' 启用每个恢复场景（注意：“Item”属性是默认属性，且可省略）
Next

' 启用恢复机制（使用出错时的默认设置）
qtTestRecovery.Enabled = True

'确保恢复机制被设置为仅在出错时激活
qtTestRecovery.SetActivationMode "OnError"
'OnError 是默认值，另一个选项是“OnEveryStep”。


Set qtApp = Nothing ' 释放 Application 对象
Set qtTestRecovery = Nothing ' 释放 Recovery 对象
