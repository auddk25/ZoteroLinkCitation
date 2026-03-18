Public Sub ZoteroLinkCitation()
    Dim nStart&, nEnd&
    nStart = Selection.Start
    nEnd = Selection.End

    ' 关闭屏幕刷新加速运行
    Application.ScreenUpdating = False

    ' 强制将"超链接"和"已访问的超链接"样式都设为黑色无下划线
    On Error Resume Next
    ActiveDocument.Styles(wdStyleHyperlink).Font.Color = wdColorBlack
    ActiveDocument.Styles(wdStyleHyperlink).Font.Underline = wdUnderlineNone
    ActiveDocument.Styles(wdStyleHyperlinkFollowed).Font.Color = wdColorBlack
    ActiveDocument.Styles(wdStyleHyperlinkFollowed).Font.Underline = wdUnderlineNone
    On Error GoTo 0

    On Error GoTo ErrorHandler

    Dim title As String, titleAnchor As String
    Dim fieldCode As String
    Dim numOrYear As String, lnkcap As String
    Dim n1&, n2&
    Dim i&

    Dim usedBookmarks As Object
    Set usedBookmarks = CreateObject("Scripting.Dictionary")
    Dim array_RefTitle(200) As String

    ActiveWindow.View.ShowFieldCodes = True
    Selection.Find.ClearFormatting

    ' 寻找文末的 Zotero 参考文献列表
    With Selection.Find
        .Text = "^d ADDIN ZOTERO_BIBL"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute

    If Not Selection.Find.Found Then
        ActiveWindow.View.ShowFieldCodes = False
        ActiveDocument.Range(nStart, nEnd).Select
        Application.ScreenUpdating = True
        MsgBox "未找到 Zotero 参考文献列表（ZOTERO_BIBL 域）。", vbExclamation, "终止"
        Exit Sub
    End If

    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Zotero_Bibliography"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With

    ' 收集所有 Zotero 引注域索引
    Dim aField As Field
    Dim zotFieldCount As Long
    Dim zotFieldIndices() As Long
    zotFieldCount = 0
    Dim fi As Long
    For fi = 1 To ActiveDocument.Fields.Count
        If InStr(ActiveDocument.Fields(fi).Code, "ADDIN ZOTERO_ITEM") > 0 Then
            zotFieldCount = zotFieldCount + 1
            ReDim Preserve zotFieldIndices(1 To zotFieldCount)
            zotFieldIndices(zotFieldCount) = fi
        End If
    Next fi

    If zotFieldCount = 0 Then
        ActiveWindow.View.ShowFieldCodes = False
        ActiveDocument.Range(nStart, nEnd).Select
        Application.ScreenUpdating = True
        MsgBox "未找到 Zotero 引注域。", vbExclamation, "终止"
        Exit Sub
    End If

    Dim failedCitations As String
    failedCitations = ""
    Dim Titles_in_Cit As Long

    ' 倒序遍历
    Dim zi As Long
    For zi = zotFieldCount To 1 Step -1
        On Error GoTo CitationError
        Set aField = ActiveDocument.Fields(zotFieldIndices(zi))
        fieldCode = aField.Code

            Dim plCitStrBeg As String
            plCitStrBeg = """plainCitation"""
            n1 = InStr(fieldCode, plCitStrBeg)

            If n1 > 0 Then
                ' 清除已有超链接
                Dim hl As Long
                For hl = aField.Result.Hyperlinks.Count To 1 Step -1
                    aField.Result.Hyperlinks(hl).Delete
                Next hl

                ' 清空标题数组
                Erase array_RefTitle

                ' ========== 提取所有 title ==========
                Dim workCode As String
                workCode = Replace(fieldCode, "\""", Chr(1))
                i = 0
                Dim searchPos As Long
                searchPos = 1
                Do While searchPos <= Len(workCode)
                    Dim foundPos As Long
                    foundPos = FindExactTitleKey(workCode, searchPos)
                    If foundPos = 0 Then Exit Do

                    n1 = foundPos + Len("""title"":""")
                    n2 = InStr(Mid(workCode, n1, Len(workCode) - n1), """,""") - 1 + n1
                    If n2 < n1 Then
                        n2 = InStr(Mid(workCode, n1, Len(workCode) - n1), "}") - 1 + n1 - 1
                    End If
                    If n2 < n1 Then n2 = n1

                    Dim rawTitle As String
                    rawTitle = DecodeUnicodeEscapes(Replace(Mid(workCode, n1, n2 - n1), Chr(1), """"))
                    array_RefTitle(i) = StripHtmlTags(rawTitle)
                    searchPos = n2 + 1
                    i = i + 1
                    If i > 199 Then Exit Do
                Loop
                Titles_in_Cit = i

                ' ========== 【核心重写】直接从参考文献列表获取真实编号 ==========
                ' 不再依赖 plainCitation 解析 RefNumber，
                ' 而是：找到标题在参考文献中的位置 → 读取 [N] → 用 N 在引文域中搜索
                Dim t As Long
                For t = 0 To Titles_in_Cit - 1
                    title = array_RefTitle(t)
                    If title = "" Then GoTo NextTitle
                    titleAnchor = MakeValidBMName(title)

                    ' 防止书签名冲突
                    Dim baseBMName As String
                    Dim bmSuffix As Long
                    baseBMName = titleAnchor
                    bmSuffix = 2
                    Do While usedBookmarks.Exists(titleAnchor)
                        titleAnchor = Left(baseBMName, 40 - Len(CStr(bmSuffix))) & bmSuffix
                        bmSuffix = bmSuffix + 1
                    Loop
                    usedBookmarks(titleAnchor) = True

                    ' --- 在参考文献列表中查找标题 ---
                    ActiveWindow.View.ShowFieldCodes = False
                    Selection.GoTo What:=wdGoToBookmark, Name:="Zotero_Bibliography"
                    Selection.Collapse Direction:=wdCollapseStart

                    Selection.Find.ClearFormatting
                    With Selection.Find
                        .Text = Left(title, 255)
                        .Replacement.Text = ""
                        .Forward = True
                        .Wrap = wdFindStop
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                    End With
                    Selection.Find.Execute

                    Dim titleVerified As Boolean
                    titleVerified = Selection.Find.Found

                    ' 长标题验证
                    If titleVerified And Len(title) > 255 Then
                        Dim verifyEnd As Long
                        verifyEnd = Selection.Start + Len(title)
                        If verifyEnd <= ActiveDocument.Content.End Then
                            Dim verifyRange As Range
                            Set verifyRange = ActiveDocument.Range(Selection.Start, verifyEnd)
                            If verifyRange.Text <> title Then
                                titleVerified = False
                            End If
                        End If
                    End If

                    ' 范围验证
                    If titleVerified And ActiveDocument.Bookmarks.Exists("Zotero_Bibliography") Then
                        Dim biblRange As Range
                        Set biblRange = ActiveDocument.Bookmarks("Zotero_Bibliography").Range
                        If Selection.Start < biblRange.Start Or Selection.End > biblRange.End Then
                            titleVerified = False
                        End If
                    End If

                    If titleVerified Then
                        ' 扩展选区到整个参考文献条目
                        Selection.MoveStartUntil ("["), Count:=wdBackward
                        Selection.MoveEndUntil (vbCr)
                        Dim entryText As String
                        entryText = Selection.Text
                        lnkcap = Left(entryText, 70)

                        ' 【核心】从参考文献条目开头的 [N] 中提取真实编号 N
                        Dim actualRefNum As String
                        actualRefNum = ""
                        Dim bracketEnd As Long
                        bracketEnd = InStr(entryText, "]")
                        If Left(entryText, 1) = "[" And bracketEnd > 2 Then
                            actualRefNum = Mid(entryText, 2, bracketEnd - 2)
                        End If

                        ' 创建书签
                        Selection.Shrink
                        With ActiveDocument.Bookmarks
                            .Add Range:=Selection.Range, Name:=titleAnchor
                            .DefaultSorting = wdSortByName
                            .ShowHidden = True
                        End With

                        If Not ActiveDocument.Bookmarks.Exists(titleAnchor) Then GoTo NextTitle

                        If actualRefNum = "" Then GoTo NextTitle

                        ' --- 用真实编号在引文域结果中搜索 ---
                        Dim searchFound As Boolean
                        searchFound = False

                        Dim dashVariants(2) As String
                        dashVariants(0) = actualRefNum
                        dashVariants(1) = Replace(Replace(actualRefNum, ChrW(8211), "-"), ChrW(8212), "-")
                        dashVariants(2) = Replace(Replace(actualRefNum, "-", ChrW(8211)), ChrW(8212), ChrW(8211))

                        Dim v As Long
                        For v = 0 To 2
                            If searchFound Then Exit For

                            aField.Select
                            Selection.Find.ClearFormatting
                            With Selection.Find
                                .Text = dashVariants(v)
                                .Replacement.Text = ""
                                .Forward = True
                                .Wrap = wdFindStop
                                .Format = False
                                .MatchCase = False
                                .MatchWholeWord = False
                            End With
                            Selection.Find.Execute

                            If Selection.Find.Found Then
                                If Selection.Start >= aField.Result.Start And Selection.End <= aField.Result.End Then
                                    numOrYear = Selection.Range.Text & ""
                                    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, _
                                        Address:="", SubAddress:=titleAnchor, _
                                        ScreenTip:=lnkcap, TextToDisplay:="" & numOrYear
                                    searchFound = True
                                End If
                            End If
                        Next v

                        ' 兜底：给整个引注块加链接
                        If Not searchFound Then
                            ActiveDocument.Hyperlinks.Add Anchor:=aField.Result, _
                                Address:="", SubAddress:=titleAnchor, ScreenTip:=lnkcap
                        End If

                        ' 去除蓝色下划线
                        aField.Select
                        With Selection.Font
                             .Underline = wdUnderlineNone
                             .ColorIndex = wdBlack
                        End With

                    Else
                        Dim failReason As String
                        If Not Selection.Find.Found Then
                            failReason = "Find未找到"
                        Else
                            failReason = "范围检查失败"
                        End If
                        failedCitations = failedCitations & Left(title, 30) & " | " & failReason & vbCrLf
                    End If
NextTitle:
                Next t
            End If
        GoTo NextCitation
CitationError:
        On Error Resume Next
        aField.Select
        Selection.Font.Underline = wdUnderlineNone
        Selection.Font.ColorIndex = wdBlack
        On Error GoTo 0
        Resume NextCitation
NextCitation:
    Next zi
    On Error GoTo ErrorHandler

    ActiveWindow.View.ShowFieldCodes = False
    ActiveDocument.Range(nStart, nEnd).Select
    Application.ScreenUpdating = True

    If failedCitations <> "" Then
        MsgBox "引注超链接已生成，但以下标题未能匹配：" & vbCrLf & vbCrLf & Left(failedCitations, 800), vbExclamation, "部分完成"
    Else
        MsgBox "引注超链接已全部生成完毕！", vbInformation, "完成"
    End If

    Exit Sub

ErrorHandler:
    ActiveWindow.View.ShowFieldCodes = False
    Application.ScreenUpdating = True

End Sub

Function MakeValidBMName(ByVal strIn As String)
    Dim pFirstChr As String
    Dim i As Long
    Dim tempStr As String
    strIn = Trim(strIn)
    pFirstChr = Left(strIn, 1)
    If Not pFirstChr Like "[A-Za-z]" Then
        strIn = "A_" & strIn
    End If
    For i = 1 To Len(strIn)
        Select Case AscW(Mid$(strIn, i, 1))
        Case 48 To 57, 65 To 90, 97 To 122
            tempStr = tempStr & Mid$(strIn, i, 1)
        Case Else
            tempStr = tempStr & "_"
        End Select
    Next i
    Do While InStr(tempStr, "__") > 0
        tempStr = Replace(tempStr, "__", "_")
    Loop
    MakeValidBMName = Left(tempStr, 40)
End Function

Function DecodeUnicodeEscapes(ByVal s As String) As String
    Dim pos As Long
    Dim codeStr As String
    Dim codeVal As Long
    pos = InStr(s, "\u")
    Do While pos > 0
        If pos + 5 <= Len(s) Then
            codeStr = Mid(s, pos + 2, 4)
            On Error Resume Next
            codeVal = CLng("&H" & codeStr)
            If Err.Number = 0 Then
                s = Left(s, pos - 1) & ChrW(codeVal) & Mid(s, pos + 6)
            Else
                Err.Clear
                pos = pos + 2
            End If
            On Error GoTo 0
        Else
            Exit Do
        End If
        pos = InStr(pos, s, "\u")
    Loop
    s = Replace(s, "\-", "-")
    DecodeUnicodeEscapes = s
End Function

Function FindExactTitleKey(ByVal s As String, ByVal startPos As Long) As Long
    Dim target As String
    target = """title"":"""
    Dim pos As Long
    pos = InStr(startPos, s, target)
    Do While pos > 0
        If pos = 1 Then
            FindExactTitleKey = pos
            Exit Function
        End If
        Dim prevChar As String
        prevChar = Mid(s, pos - 1, 1)
        If prevChar = "," Or prevChar = "{" Or prevChar = " " Or prevChar = vbTab Or prevChar = vbLf Or prevChar = vbCr Then
            FindExactTitleKey = pos
            Exit Function
        End If
        pos = InStr(pos + 1, s, target)
    Loop
    FindExactTitleKey = 0
End Function

Function StripHtmlTags(ByVal s As String) As String
    Dim result As String
    Dim inTag As Boolean
    Dim ch As String
    Dim i As Long
    result = ""
    inTag = False
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch = "<" Then
            inTag = True
        ElseIf ch = ">" Then
            inTag = False
        ElseIf Not inTag Then
            result = result & ch
        End If
    Next i
    StripHtmlTags = result
End Function
