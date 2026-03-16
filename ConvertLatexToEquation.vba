Sub ConvertLatexToEquation()
    ' 关闭屏幕更新，大幅提升长文档的转换速度，避免屏幕闪烁
    Application.ScreenUpdating = False 
    
    Dim rng As Range
    
    ' ==========================================
    ' 第一阶段：优先处理 $$ ... $$ (行间公式/独立公式)
    ' 必须先处理双$，否则单$的逻辑会将其拆散
    ' ==========================================
    Set rng = ActiveDocument.Range
    With rng.Find
        .ClearFormatting
        .Text = "\$\$[!\$]{1,}\$\$"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            Dim innerText As String
            ' 提取中间的文本，去除前后各2个字符的 $$
            innerText = Mid(rng.Text, 3, Len(rng.Text) - 4)
            
            ' 【关键修复】：$$ 公式内部经常会有换行。
            ' 将硬回车(vbCr)替换为软换行(Chr(11))，保证 Word 公式将其视为同一个整体框，而不会被腰斩
            innerText = Replace(innerText, vbCr, Chr(11))
            
            rng.Text = innerText
            rng.End = rng.Start + Len(innerText)
            
            Dim mathRange As Range
            Set mathRange = ActiveDocument.OMaths.Add(rng)
            
            If mathRange.OMaths.Count > 0 Then
                ' 将 $$ 公式强制设置为“独立显示模式” (居中独占一行)
                mathRange.OMaths(1).Type = wdOMathDisplay
                mathRange.OMaths(1).BuildUp
            End If
            
            rng.Collapse wdCollapseEnd
        Loop
    End With

    ' ==========================================
    ' 第二阶段：处理剩下的 $ ... $ (行内公式/内嵌公式)
    ' ==========================================
    Set rng = ActiveDocument.Range
    With rng.Find
        .ClearFormatting
        .Text = "\$[!\$]{1,}\$"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            ' 奇偶错位保护：行内单 $ 公式通常不会跨越段落。如果包含回车，说明可能是漏掉了 $
            If InStr(rng.Text, vbCr) > 0 Then
                ' 只往后移动1个字符，解救被错配的无辜文本
                rng.SetRange rng.Start + 1, rng.Start + 1
                GoTo ContinueInlineLoop
            End If
            
            Dim innerText2 As String
            ' 提取中间的文本，去除前后各1个字符的 $
            innerText2 = Mid(rng.Text, 2, Len(rng.Text) - 2)
            
            rng.Text = innerText2
            rng.End = rng.Start + Len(innerText2)
            
            Dim mathRange2 As Range
            Set mathRange2 = ActiveDocument.OMaths.Add(rng)
            
            If mathRange2.OMaths.Count > 0 Then
                ' 将 $ 公式强制设置为“内嵌模式” (与普通文字紧凑同行)
                mathRange2.OMaths(1).Type = wdOMathInline
                mathRange2.OMaths(1).BuildUp
            End If
            
            rng.Collapse wdCollapseEnd
ContinueInlineLoop:
        Loop
    End With
    
    ' 恢复屏幕更新
    Application.ScreenUpdating = True
    MsgBox "LaTeX 公式转换完成！(V4 支持 $$ 行间公式版)", vbInformation, "转换完毕"
End Sub
