Option Explicit

' ============================================================
' Fix_Spacing_And_Quote_Font v5
'
' v4→v5 核心策略变化：占位符保护
'
'   原理：
'     1. 检测英文图表标题段落
'     2. 把标题中所有空白字符替换为占位符 ¤ (ChrW(164))
'     3. 执行所有全局清理操作（不可能删掉 ¤）
'     4. 把 ¤ 恢复为普通空格 Chr(32)
'
'   为什么这一定能工作：
'     - 宏中没有任何操作会删除 ¤ 字符
'     - 不依赖"跳过段落"或"保护检测"等脆弱逻辑
'     - 处理完成后弹窗显示第一个英文标题的处理前/后对比
'     - 如果对比显示空格还在 → 问题在 Word 渲染，不在宏
'     - 如果对比显示空格没了 → 请截图发给我
' ============================================================

' 占位符：¤ (ChrW(164))，货币符号，不可能出现在地质学论文图表标题中
Private Const SPACE_HOLDER As String = ChrW(164)

' ==================== 主入口 ====================

Sub Fix_Spacing_And_Quote_Font()
    Dim rng As Range
    Dim p As Paragraph
    Dim captionCount As Long
    Dim pRng As Range
    Dim beforeText As String
    Dim afterText As String
    Dim bmName As String
    Dim i As Long

    Set rng = GetWorkRange()
    Application.ScreenUpdating = False

    ' ====== Step 0: 检测英文标题 → 替换空格为占位符 ======
    captionCount = 0
    For Each p In rng.Paragraphs
        If IsEnglishCaptionContent(p) Then
            captionCount = captionCount + 1
            ' 用书签标记段落位置（后续恢复用）
            bmName = "_CapSave" & captionCount
            ActiveDocument.Bookmarks.Add bmName, p.Range.Duplicate
            ' 记录第一个标题的处理前文本
            If captionCount = 1 Then beforeText = CleanParaText(p.Range.Text)
            ' 替换所有空白为占位符
            ProtectSpaces p.Range.Duplicate
        End If
    Next p

    ' ====== Step 1: 关闭中日韩自动间距（全局）======
    For Each p In rng.Paragraphs
        On Error Resume Next
        p.Range.ParagraphFormat.AddSpaceBetweenFarEastAndAlpha = False
        p.Range.ParagraphFormat.AddSpaceBetweenFarEastAndDigit = False
        On Error GoTo 0
    Next p

    ' ====== Step 2: 全局空格清理 ======
    Call RemoveWeirdSpaces(rng)

    ' ====== Step 3: 引号字体 ======
    Call SetQuoteFontSongti(rng)

    ' ====== Step 4: 恢复英文标题（占位符 → 普通空格）======
    For i = 1 To captionCount
        bmName = "_CapSave" & i
        If ActiveDocument.Bookmarks.Exists(bmName) Then
            Set pRng = ActiveDocument.Bookmarks(bmName).Range.Duplicate
            RestoreSpaces pRng
            ActiveDocument.Bookmarks(bmName).Delete
            If i = 1 Then afterText = CleanParaText(pRng.Text)
        End If
    Next i

    Application.ScreenUpdating = True

    ' ====== 报告结果 ======
    Dim msg As String
    msg = "[v5] 处理完成。已保护 " & captionCount & " 个英文图表标题。"
    If captionCount > 0 And Len(beforeText) > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "▼ 第一个英文标题对比" & vbCrLf & _
              "【前】" & beforeText & vbCrLf & _
              "【后】" & afterText
    End If
    MsgBox msg, vbInformation
End Sub

' ==================== 诊断工具 ====================

Sub DiagnoseCaptionStyle()
    Dim p As Paragraph
    Dim rng As Range
    Dim txt As String
    Dim msg As String
    Dim i As Long
    Dim chCode As Long
    Dim nbspCount As Long, spaceCount As Long, otherCount As Long

    Set rng = Selection.Range
    Set p = rng.Paragraphs(1)
    txt = p.Range.Text

    nbspCount = 0: spaceCount = 0: otherCount = 0

    For i = 1 To Len(txt)
        chCode = AscW(Mid(txt, i, 1))
        Select Case chCode
            Case 32: spaceCount = spaceCount + 1
            Case 160: nbspCount = nbspCount + 1
            Case 8194, 8195, 8196, 8197, 8198, 8199, 8200, 8201, 8202, 8239, 8287, 12288
                otherCount = otherCount + 1
        End Select
    Next i

    msg = "========== 段落样式 ==========" & vbCrLf
    On Error Resume Next
    msg = msg & "段落样式: " & p.Style.NameLocal & vbCrLf
    If p.Range.Characters.Count > 1 Then
        msg = msg & "首字符样式: " & p.Range.Characters(1).Style.NameLocal & vbCrLf
    End If
    On Error GoTo 0

    msg = msg & vbCrLf & "========== 空格统计 ==========" & vbCrLf
    msg = msg & "普通空格 Chr(32): " & spaceCount & vbCrLf
    msg = msg & "NBSP ChrW(160):   " & nbspCount & vbCrLf
    msg = msg & "其他特殊空格:      " & otherCount & vbCrLf

    msg = msg & vbCrLf & "========== 保护检测 ==========" & vbCrLf
    msg = msg & "IsEnglishCaptionContent: " & IIf(IsEnglishCaptionContent(p), "True", "False") & vbCrLf

    msg = msg & vbCrLf & "========== 前30字符码 ==========" & vbCrLf
    Dim maxC As Long
    maxC = 30
    If Len(txt) < maxC Then maxC = Len(txt)
    For i = 1 To maxC
        chCode = AscW(Mid(txt, i, 1))
        msg = msg & i & ": ChrW(" & chCode & ") "
        Select Case chCode
            Case 32: msg = msg & "[空格]"
            Case 13: msg = msg & "[回车]"
            Case 160: msg = msg & "[NBSP]"
            Case 8194: msg = msg & "[En空格]"
            Case 8195: msg = msg & "[Em空格]"
            Case 8201: msg = msg & "[细空格]"
            Case 8239: msg = msg & "[窄NBSP]"
            Case 12288: msg = msg & "[全角空格]"
            Case Else: msg = msg & "[" & Mid(txt, i, 1) & "]"
        End Select
        msg = msg & vbCrLf
    Next i

    MsgBox msg, vbInformation, "诊断 v5"
End Sub

' ==================== 辅助函数 ====================

Function GetWorkRange() As Range
    If Selection.Range.Start <> Selection.Range.End Then
        Set GetWorkRange = Selection.Range.Duplicate
    Else
        Set GetWorkRange = ActiveDocument.Content
    End If
End Function

Function CleanParaText(ByVal txt As String) As String
    If Len(txt) > 0 And Right(txt, 1) = Chr(13) Then
        CleanParaText = Left(txt, Len(txt) - 1)
    Else
        CleanParaText = txt
    End If
End Function

' ==================== 英文标题检测 ====================

Function IsEnglishCaptionContent(ByVal p As Paragraph) As Boolean
    Dim txt As String
    Dim chCode As Long

    IsEnglishCaptionContent = False
    txt = p.Range.Text

    ' 去掉段落标记
    If Len(txt) > 0 And Right(txt, 1) = Chr(13) Then
        txt = Left(txt, Len(txt) - 1)
    End If

    ' 去掉所有前导空白（含占位符 ¤，防止已保护的段落被误判）
    Do While Len(txt) > 0
        chCode = AscW(Left(txt, 1))
        If IsWhitespaceOrHolder(chCode) Then
            txt = Mid(txt, 2)
        Else
            Exit Do
        End If
    Loop

    ' 宽松匹配：以 "Fig" 或 "Table" 开头
    If Left(txt, 3) = "Fig" Or Left(txt, 5) = "Table" Then
        IsEnglishCaptionContent = True
    End If
End Function

Function IsWhitespaceOrHolder(ByVal chCode As Long) As Boolean
    Select Case chCode
        Case 9, 10, 13, 32, 160, 8194, 8195, 8196, 8197, 8198, _
             8199, 8200, 8201, 8202, 8232, 8233, 8239, 8287, 12288, _
             164  ' 占位符 ¤
            IsWhitespaceOrHolder = True
        Case Else
            IsWhitespaceOrHolder = False
    End Select
End Function

' ==================== 占位符保护/恢复 ====================

Sub ProtectSpaces(ByVal rng As Range)
    ' 把所有可能的空白字符替换为占位符 ¤
    ReplaceLiteralInRange rng, Chr(32), SPACE_HOLDER       ' 普通空格
    ReplaceLiteralInRange rng, ChrW(160), SPACE_HOLDER      ' NBSP
    ReplaceLiteralInRange rng, ChrW(8192), SPACE_HOLDER     ' En Quad
    ReplaceLiteralInRange rng, ChrW(8193), SPACE_HOLDER     ' Em Quad
    ReplaceLiteralInRange rng, ChrW(8194), SPACE_HOLDER     ' En Space
    ReplaceLiteralInRange rng, ChrW(8195), SPACE_HOLDER     ' Em Space
    ReplaceLiteralInRange rng, ChrW(8196), SPACE_HOLDER     ' Three-Per-Em
    ReplaceLiteralInRange rng, ChrW(8197), SPACE_HOLDER     ' Four-Per-Em
    ReplaceLiteralInRange rng, ChrW(8198), SPACE_HOLDER     ' Six-Per-Em
    ReplaceLiteralInRange rng, ChrW(8199), SPACE_HOLDER     ' Figure Space
    ReplaceLiteralInRange rng, ChrW(8200), SPACE_HOLDER     ' Punctuation Space
    ReplaceLiteralInRange rng, ChrW(8201), SPACE_HOLDER     ' Thin Space
    ReplaceLiteralInRange rng, ChrW(8202), SPACE_HOLDER     ' Hair Space
    ReplaceLiteralInRange rng, ChrW(8239), SPACE_HOLDER     ' Narrow NBSP
    ReplaceLiteralInRange rng, ChrW(8287), SPACE_HOLDER     ' Medium Math Space
    ReplaceLiteralInRange rng, ChrW(12288), SPACE_HOLDER    ' 全角空格
    ReplaceLiteralInRange rng, vbTab, SPACE_HOLDER           ' Tab
End Sub

Sub RestoreSpaces(ByVal rng As Range)
    ' 把占位符 ¤ 恢复为普通空格
    ReplaceLiteralInRange rng, SPACE_HOLDER, Chr(32)
End Sub

' ==================== 全局空格清理 ====================

Sub RemoveWeirdSpaces(ByVal rng As Range)
    Dim i As Long

    ' 特殊空格字符
    ReplaceLiteralInRange rng, ChrW(160), ""      ' NBSP
    ReplaceLiteralInRange rng, ChrW(12288), ""    ' 全角空格
    ReplaceLiteralInRange rng, ChrW(9675), ""     ' ○
    ReplaceLiteralInRange rng, ChrW(183), ""      ' ·
    ReplaceLiteralInRange rng, ChrW(8226), ""     ' •
    ReplaceLiteralInRange rng, vbTab, ""

    ' 中文图表式编号空格
    For i = 0 To 300
        ReplaceLiteralInRange rng, "图 " & CStr(i), "图" & CStr(i)
        ReplaceLiteralInRange rng, "表 " & CStr(i), "表" & CStr(i)
        ReplaceLiteralInRange rng, "式 " & CStr(i), "式" & CStr(i)
    Next i

    ' 常见中文写法
    ReplaceLiteralInRange rng, "见 图", "见图"
    ReplaceLiteralInRange rng, "见 表", "见表"
    ReplaceLiteralInRange rng, "见 式", "见式"
    ReplaceLiteralInRange rng, "(见 图", "(见图"
    ReplaceLiteralInRange rng, "（见 图", "（见图"

    ' 引用前后空隙
    ReplaceLiteralInRange rng, "[ ", "["
    ReplaceLiteralInRange rng, " ]", "]"
    ReplaceLiteralInRange rng, " [", "["
    ReplaceLiteralInRange rng, "] ", "]"

    ' 反复压缩普通空格
    For i = 1 To 5
        ReplaceLiteralInRange rng, "  ", " "
    Next i
End Sub

' ==================== 引号字体设置 ====================

Sub SetQuoteFontSongti(ByVal rng As Range)
    SetOneCharFont rng, ChrW(8220), "宋体"   ' "
    SetOneCharFont rng, ChrW(8221), "宋体"   ' "
    SetOneCharFont rng, ChrW(8216), "宋体"   ' '
    SetOneCharFont rng, ChrW(8217), "宋体"   ' '
End Sub

' ==================== 底层工具函数 ====================

Sub SetOneCharFont(ByVal rng As Range, ByVal targetChar As String, ByVal fontName As String)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = targetChar
        .Replacement.Text = targetChar
        .Replacement.Font.NameFarEast = fontName
        .Replacement.Font.Name = fontName
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub ReplaceLiteralInRange(ByVal rng As Range, ByVal findText As String, ByVal replaceText As String)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub
