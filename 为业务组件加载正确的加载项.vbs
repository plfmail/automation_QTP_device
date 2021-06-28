'************************************************************************************************************************
'描述：
'
' 该示例查找与特定业务组件关联的加载项。
' 如果尚未加载某些必需的加载项，则其加载它们，重新启动 QuickTest，打开业务组件，
' 并确认打开的文档是否确实是业务组件。
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

    arrBCAddins = qtApp.GetAssociatedAddinsForBC("[QualityCenter] Components\MyFolder\MyBC")

    ' 检查是否已加载所有必需的加载项
    blnNeedChangeAddins = False ' 假定无需作任何更改
    For Each bcAddin In arrBCAddins ' 循环与组件关联的加载项列表
        If qtApp.Addins(bcAddin).Status <> "Active" Then ' 如果存在未加载的关联加载项
            blnNeedChangeAddins = True ' 指明需要对加载的加载项进行更改
            Exit For ' 退出循环
        End If
    Next

    If qtApp.Launched And blnNeedChangeAddins Then
        qtApp.Quit ' 如果有必要进行更改，则退出 QuickTest，修改已加载的加载项
    End If

    If blnNeedChangeAddins Then
        Dim blnActivateOK
        blnActivateOK = qtApp.SetActiveAddins(arrBCAddins, errorDescription) ' 加载与组件关联的加载项并检查它们是否已加载成功。
        If Not blnActivateOK Then ' 如果在加载加载项时发生问题
            MsgBox errorDescription ' 显示包含错误的消息
            WScript.Quit ' 并结束自动程序。
        End If
    End If
End If

If Not qtApp.Launched Then ' 如果尚未打开 QuickTest
    qtApp.Launch ' 启动 QuickTest（已加载正确的加载项）
    qtApp.Visible = True ' 使 QuickTest 应用程序可见
    qtApp.TDConnection.Connect "http://qcserver/qcbin", _
              "MY_DOMAIN", "My_Project", "James", "not4you", False ' 连接到 Quality Center
End If

If qtApp.TDConnection.IsConnected Then ' 如果连接成功
    qtApp.OpenBusinessComponent "[QualityCenter] Components\MyFolder\MyBC", False ' 打开业务组件
    MsgBox qtApp.CurrentDocumentType '确认打开的文档是否是业务组件
End If
