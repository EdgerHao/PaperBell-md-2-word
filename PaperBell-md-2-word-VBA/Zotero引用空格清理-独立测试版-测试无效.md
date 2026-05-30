```vba
Option Explicit

' =======================================================
' Zotero 引用空格清理 - 独立测试版
' 用途: 清除 Zotero 引文域内部首尾空格 + 域与正文之间的多余空格
' =======================================================

Sub ZoteroSpaceCleanup()
    Dim fld As Field
    Dim rng As Range
    Dim i As Long
    Dim cleanCount As Long

    Application.ScreenUpdating = False
    cleanCount = 0

    ' 从后往前遍历，避免删除导致位置偏移
    For i = ActiveDocument.Fields.Count To 1 Step -1
        Set fld = ActiveDocument.Fields(i)

        ' 只处理 Zotero 引文域
        If InStr(1, Trim$(fld.Code.Text), "ADDIN ZOTERO_ITEM CSL_CITATION", vbTextCompare) = 1 Then

            ' === Part A：清理 field result 内部首尾空格 ===
            Set rng = fld.Result.Duplicate

            ' 去尾部空格（普通空格、NBSP、全角空格、制表符）
            Do While rng.Characters.Count > 0
                If IsSpaceChar(rng.Characters.Last.Text) Then
                    rng.Characters.Last.Delete
                    cleanCount = cleanCount + 1
                Else
                    Exit Do
                End If
            Loop

            ' 刷新 range 后去首部空格
            Set rng = fld.Result.Duplicate
            Do While rng.Characters.Count > 0
                If IsSpaceChar(rng.Characters.First.Text) Then
                    rng.Characters.First.Delete
                    cleanCount = cleanCount + 1
                Else
                    Exit Do
                End If
            Loop

            ' === Part B：清理 field 外部紧邻空格（引用域前）===
            Set rng = ActiveDocument.Range(fld.Result.Start, fld.Result.End)
            If rng.Start > ActiveDocument.Content.Start Then
                Dim rngPre As Range
                Do While rng.Start > ActiveDocument.Content.Start
                    Set rngPre = ActiveDocument.Range(rng.Start - 1, rng.Start)
                    If IsSpaceChar(rngPre.Text) Then
                        rngPre.Delete
                        cleanCount = cleanCount + 1
                    Else
                        Exit Do
                    End If
                Loop
            End If

            ' === Part C：清理 field 外部紧邻空格（引用域后）===
            Set rng = ActiveDocument.Range(fld.Result.Start, fld.Result.End)
            If rng.End < ActiveDocument.Content.End Then
                Dim rngPost As Range
                Do While rng.End < ActiveDocument.Content.End
                    Set rngPost = ActiveDocument.Range(rng.End, rng.End + 1)
                    If IsSpaceChar(rngPost.Text) Then
                        rngPost.Delete
                        cleanCount = cleanCount + 1
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
    Next i

    Application.ScreenUpdating = True
    MsgBox "Zotero 引用空格清理完成！" & vbCrLf & vbCrLf & _
           "共清理 " & cleanCount & " 处空格。", vbInformation, "完成"
End Sub

' --- 辅助函数：判断是否为空格字符 ---
Function IsSpaceChar(ByVal ch As String) As Boolean
    Select Case ch
        Case " ", ChrW(160), ChrW(12288), vbTab
            IsSpaceChar = True
        Case Else
            IsSpaceChar = False
    End Select
End Function
```
