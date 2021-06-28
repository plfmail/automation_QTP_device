
'************************************************************************************************************************
'描述：
'
'该示例打开 QuickTest，并配置 WinList 对象类的对象标识设置。
'************************************************************************************************************************

Dim qtApp ' As QuickTest.Application ' 声明 Application 对象变量
Dim qtIdent 'As QuickTest.ObjectIdentification ' 声明 ObjectIdentification 对象变量
Dim qtWinListIdent 'As QuickTest.TestObjectClassIdentification ' 声明 WinList 对象类标识的变量
Dim intPosition ' 声明变量，用于存储位置

' 打开 QuickTest 并设置变量
Set qtApp = CreateObject("QuickTest.Application") ' 创建 Application 对象
qtApp.Launch ' 启动 QuickTest
qtApp.Visible = True ' 使 QuickTest 应用程序可见

Set qtIdent = qtApp.Options.ObjectIdentification ' 返回 ObjectIdentification 对象
Set qtWinListIdent = qtIdent.Item("WinList") ' 返回 WinList 对象类的对象标识属性的集合

qtIdent.ResetAll ' 将 WinList 对象的对象标识描述重置为默认属性设置
qtWinListIdent.OrdinalIdentifier = "Index" ' 将 Index 设置为顺序标识符

' 配置强制属性
intPosition = qtWinListIdent.MandatoryProperties.Find("nativeclass") ' 查找“nativeclass”强制属性的位置
qtWinListIdent.MandatoryProperties.Remove intPosition ' 从列表中删除“nativeclass”强制属性
If qtWinListIdent.AvailableProperties.Find("items count") <> -1 Then ' 如果“items count”是可用的 WinList 属性
    qtWinListIdent.MandatoryProperties.Add "items count" ' 将其添加为强制属性
End If

' 配置辅助属性
qtWinListIdent.AssistiveProperties.RemoveAll ' 删除所有辅助属性
qtWinListIdent.AssistiveProperties.Add "all items" ' 将“all items”属性添加为辅助属性
qtWinListIdent.AssistiveProperties.Add "width", 1 ' 将“width”添加为第一个辅助属性
qtWinListIdent.AssistiveProperties.Add "height", -1 ' 将“height”添加为最后一个辅助属性
qtWinListIdent.AssistiveProperties.MoveToPos 2, 1 ' 将第二个辅助属性（当前为“all items”）移动到列表的第一个位置
' 配置智能标识
qtWinListIdent.EnableSmartIdentification = True ' 启用 WinList 对象的智能标识机制
If qtWinListIdent.BaseFilterProperties.Count = 0 Then ' 如果不存在基本筛选器属性
    qtWinListIdent.BaseFilterProperties.Add "x" ' 将“x”属性添加为基本筛选器属性
    qtWinListIdent.BaseFilterProperties.Add "y" ' 将“y”属性添加为基本筛选器属性
End If
qtWinListIdent.OptionalFilterProperties.Add "abs_x", 1 ' 将“abs_x”添加为第一个可选筛选器属性
qtWinListIdent.OptionalFilterProperties.Add "abs_y", 2 ' 将“abs_y”添加为第二个可选筛选器属性

Set qtWinListIdent = Nothing ' 释放 WinList 标识对象
Set qtIdent = Nothing ' 释放 ObjectIdentification 对象
Set qtApp = Nothing ' 释放 Application 对象
