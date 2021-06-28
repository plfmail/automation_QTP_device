
'************************************************************************************************************************
'描述：
'
' 该示例为特定的业务流程测试找到关联的加载项的列表。
' 然后新建业务组件，并与关联 BPT 的业务组件
'关联相同的加载项。
'
'************************************************************************************************************************


Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim arrBCAddins ' 声明变量，用于存储与组件关联的加载项
Dim blnNeedChangeAddins ' 声明一个标志，用于指明当前是否已加载与组件关联的加载项

Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见

qtApp.TDConnection.Connect "http://qcserver/qcbin", _
              "MY_DOMAIN", "My_Project", "James", "not4you", False ' 连接到 Quality Center


If qtApp.TDConnection.IsConnected Then ' 如果连接成功

'找到与业务流程测试关联的加载项，并关联到具有新的业务组件的
'同一列表。
    arrBCAddins = qtApp.GetAssociatedAddinsForBPT("[QualityCenter] Subject\MyFolder\MyBPT")
    qtApp.NewBusinessComponent
    qtApp.BusinessComponent.SetAssociatedAddins arrBCAddins

End If

