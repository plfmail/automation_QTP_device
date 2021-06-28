'************************************************************************************************************************
'描述：
'
'该示例打开具有预定义参数的测试，
' 获取参数定义的集合，并循环显示每个参数的详细信息，
' 获取运行时参数的集合，更改其中一个参数的值，运行具有参数的测试，
' 测试运行后，显示其中一个输出参数的值。
'
'假定：
' 测试 D:\Tests\Mytest 包含名为“InParam1”的输入参数和名为“OutParam1”的输出参数
'************************************************************************************************************************


Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim pDefColl 'As QuickTest.ParameterDefinitions ' 声明 Parameter Definitions 集合
Dim pDef ' As QuickTest.ParameterDefinition ' 声明 ParameterDefinition 对象
Dim rtParams 'As QuickTest.Parameters ' 声明 Parameters 集合
Dim rtParam ' As QuickTest.Parameter ' 声明 Parameter 对象
'Dim cnt, Indx As Integer

Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见

qtApp.Open "D:\Tests\MyTest"

' 检索为测试定义的参数集合。
Set pDefColl = qtApp.Test.ParameterDefinitions

cnt = pDefColl.Count
Indx = 1

' 显示集合中每个参数的名称和值。
While Indx <= cnt
    Set pDef = pDefColl.Item(Indx)
    MsgBox "Param name: " & pDef.Name & "; Type: " & pDef.Type & "; InOut: " & pDef.InOut & "; Description: " _
        & pDef.Description & "Default value: " & pDef.DefaultValue
    Indx = Indx + 1
Wend

Set rtParams = pDefColl.GetParameters() ' 检索为测试定义的参数集合。

Set rtParam = rtParams.Item("InParam1") ' 检索特定的参数。
rtParam.Value = "Hello" ' 更改参数的值。

qtApp.Test.Run , True, rtParams ' 运行已更改参数的测试。

MsgBox rtParams.Item("OutParam1").Value ' 测试运行后，显示输出参数的值。
