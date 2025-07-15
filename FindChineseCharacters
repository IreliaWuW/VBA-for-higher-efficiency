'================================================================================
' Module:  FindChineseCharacters
' Purpose: Provides tools to identify and clear cells containing Chinese characters.
'================================================================================

Option Explicit

' Main Function: Highlights all cells within a selected range that contain Chinese characters.
Sub HighlightCellsWithChineseChars()
    
    ' --- Declaration ---
    Dim TargetRange As Range      ' 选择的目标范围
    Dim Cell As Range             ' 循环的单个单元格
    Dim FoundCounter As Long      ' 记录找到的单元格数量
    Dim i As Long                 ' 字符串内字符的循环
    Dim SingleChar As String      ' 存储单个字符
    
    ' --- 初始化 ---
    FoundCounter = 0
    On Error GoTo ErrorHandler ' 简单错误处理
    
    ' --- 获取选择的范围 ---
    Set TargetRange = Application.InputBox("请用鼠标选择需要检查的单元格范围：", "选择范围", Type:=8)
    
    ' 如果点击“取消”，则退出程序
    If TargetRange Is Nothing Then
        MsgBox "操作已取消。", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False ' 关闭屏幕刷新可提速
    
    ' --- 核心逻辑：遍历每个单元格并检查字符 ---
    For Each Cell In TargetRange.Cells
        ' 仅处理非空且为文本的单元格
        If Not IsEmpty(Cell.Value) And VarType(Cell.Value) = vbString Then
            ' 遍历单元格内的每一个字符
            For i = 1 To Len(Cell.Value)
                SingleChar = Mid(Cell.Value, i, 1)
                
                ' 使用 AscW 函数检查字符的 Unicode 值
                ' 中文（CJK统一汉字）的基本Unicode范围是 &H4E00 到 &H9FFF
                If AscW(SingleChar) >= &H4E00 And AscW(SingleChar) <= &H9FFF Then
                    Cell.Interior.Color = vbYellow ' 设置背景色为黄色
                    FoundCounter = FoundCounter + 1 ' 计数器加一
                    Exit For ' 找到一个中文字符后就无需再检查此单元格，跳出内层循环
                End If
            Next i
        End If
    Next Cell
    
    Application.ScreenUpdating = True ' 恢复屏幕刷新
    
    ' --- 结果报告 ---
    If FoundCounter > 0 Then
        MsgBox "检查完成！" & vbCrLf & vbCrLf & "共找到并高亮了 " & FoundCounter & " 个含中文字符的单元格。", vbInformation, "任务完成"
    Else
        MsgBox "检查完成！" & vbCrLf & vbCrLf & "在选择的范围内未发现任何中文字符。", vbInformation, "任务完成"
    End If
    
    Exit Sub

' --- 错误处理程序 ---
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "程序运行出错：" & vbCrLf & Err.Description, vbCritical, "错误"
    
End Sub

' Helper Function: Clears the background color of all cells in a selected range.
Sub ClearHighlighting()

    Dim TargetRange As Range
    
    On Error Resume Next ' 如果取消选择，则忽略错误
    Set TargetRange = Application.InputBox("请选择需要清除高亮的单元格范围：", "选择范围", Type:=8)
    
    If TargetRange Is Nothing Then
        MsgBox "操作已取消。", vbInformation
        Exit Sub
    End If
    
    ' 将选中单元格的背景色设置为“无填充”
    TargetRange.Interior.ColorIndex = xlNone
    
    MsgBox "选定区域的高亮已清除。", vbInformation, "操作成功"

End Sub
