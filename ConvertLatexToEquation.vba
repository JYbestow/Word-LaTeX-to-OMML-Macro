Sub ConvertLatexToEquation()
    Dim rng As Range
    ' 将搜索范围设定为整个文档
    Set rng = ActiveDocument.Range

    ' 设置查找条件
    With rng.Find
        .ClearFormatting
        ' 使用通配符查找规则：匹配以 $ 开头和结尾，且中间不包含 $ 的字符串
        .Text = "\$[!\$]{1,}\$"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop

        ' 循环查找文档中所有匹配项
        Do While .Execute
            ' 去除文本前后的 $ 符号
            rng.Text = Mid(rng.Text, 2, Len(rng.Text) - 2)
            
            ' 将这段去掉了 $ 的文本转换为 Word 公式对象 (OMath)
            Dim mathRange As Range
            Set mathRange = ActiveDocument.OMaths.Add(rng)
            
            ' 构建公式（将其从纯文本转换为渲染后的专业数学公式格式）
            If mathRange.OMaths.Count > 0 Then
                mathRange.OMaths(1).BuildUp
            End If
            
            ' 将光标/搜索范围折叠到当前公式之后，继续往后查找
            rng.Collapse wdCollapseEnd
        Loop
    End With
    
    MsgBox "LaTeX 公式转换完成！", vbInformation, "转换完毕"
End Sub