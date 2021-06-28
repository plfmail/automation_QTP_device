
'************************************************************************************************************************
'描述：
'
'该示例使用与指定测试关联的加载项列表中的
' Oracle 加载项替换 Java 加载项。
'
'假定：
'
'QuickTest 中当前未打开没有保存的测试。
'有关详细信息，请参阅 Test.SaveAs 方法的示例。
'************************************************************************************************************************


Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量

Directory = InputBox("Enter the path of the test you want to convert.")
Call ConvertTest(Directory)


Private Function ConvertTest(TestPath)
    ' 创建 Application 对象
    Set qtApp = CreateObject("QuickTest.Application")
    ' 打开测试。
    Call qtApp.Open(TestPath)
    ' 检索与测试关联的加载项的列表
    Addins = qtApp.Test.GetAssociatedAddins()
    ' 如果在检索的列表中出现 Java，则用 Oracle 将其替换。
    For i = 0 To UBound(Addins)
        If Addins(i) = "Java" Then
            Addins(i) = "Oracle"
        End If
    Next
    ' 将新的关联的加载项列表应用于测试。
    If Not qtApp.Test.SetAssociatedAddins(Addins, ErrorDesc) Then
                MsgBox "Unable to modify the associated add-ins for this test: " _
                    & Chr(13) & TestPath _
                    & Chr(13) & ErrorDesc
                Exit Function
            End If
    ' 如果更改成功，则保存测试。
    Call qtApp.Test.Save
End Function
