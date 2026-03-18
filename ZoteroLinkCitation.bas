Public Sub ZoteroLinkCitation()
    Dim nStart&, nEnd&
    nStart = Selection.Start
    nEnd = Selection.End

    ' 关闭屏幕刷新加速运行
    Application.ScreenUpdating = False

    ' 【修复3】强制将"超链接"和"已访问的超链接"样式都设为黑色无下划线
    ' 这样点击后链接也不会变紫色
    ' 【审计修复】仅在样式操作处局部容错，避免全局吞掉所有错误
    On Error Resume Next
    ActiveDocument.Styles(wdStyleHyperlink).Font.Color = wdColorBlack
    ActiveDocument.Styles(wdStyleHyperlink).Font.Underline = wdUnderlineNone
    ActiveDocument.Styles(wdStyleHyperlinkFollowed).Font.Color = wdColorBlack
    ActiveDocument.Styles(wdStyleHyperlinkFollowed).Font.Underline = wdUnderlineNone
    On Error GoTo 0

    ' 【审计修复】全局错误处理：确保异常时恢复 ShowFieldCodes 和 ScreenUpdating
    On Error GoTo ErrorHandler

    Dim title As String, titleAnchor As String
    Dim fieldCode As String
    Dim numOrYear As String, lnkcap As String
    Dim n1&, n2&
    Dim Titles_in_Cit&, Refs_in_Cit&, i&, Refs&

    ' 【审计修复】用 Dictionary 追踪已用书签名，防止截断后冲突
    Dim usedBookmarks As Object
    Set usedBookmarks = CreateObject("Scripting.Dictionary")
    Dim array_RefTitle(200) As String
    Dim RefNumber(200) As String

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

    ' 【审计修复】未找到参考文献列表时，提前终止并提示用户
    If Not Selection.Find.Found Then
        ActiveWindow.View.ShowFieldCodes = False
        ActiveDocument.Range(nStart, nEnd).Select
        Application.ScreenUpdating = True
        MsgBox "未找到 Zotero 参考文献列表（ZOTERO_BIBL 域），请确认文档中已插入 Zotero 参考文献。", vbExclamation, "终止"
        Exit Sub
    End If

    ' 为参考文献列表打上大范围书签
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Zotero_Bibliography"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With

    ' 【审计修复】先收集所有 Zotero 域的索引，再倒序遍历，
    ' 避免 Hyperlinks.Add 插入新 Field 导致 For Each 迭代器异常
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

    ' 【审计修复】无 Zotero 引注域时提前终止，避免空跑后显示"成功"
    If zotFieldCount = 0 Then
        ActiveWindow.View.ShowFieldCodes = False
        ActiveDocument.Range(nStart, nEnd).Select
        Application.ScreenUpdating = True
        MsgBox "未找到 Zotero 引注域（ZOTERO_ITEM），请确认文档中已插入 Zotero 引注。", vbExclamation, "终止"
        Exit Sub
    End If

    ' 用于调试：记录失败的引文
    Dim failedCitations As String
    failedCitations = ""

    ' 倒序遍历：新插入的 Field 索引在后方，不影响前方未处理的元素
    Dim zi As Long
    For zi = zotFieldCount To 1 Step -1
        ' 【兜底】单条引注出错不中断整个宏，跳过继续处理下一条
        On Error GoTo CitationError
        Set aField = ActiveDocument.Fields(zotFieldIndices(zi))
        fieldCode = aField.Code

            Dim plain_Cit As String
            Dim plCitStrBeg As String, plCitStrEnd As String
            plCitStrBeg = """plainCitation"":""["
            plCitStrEnd = "]"""
            n1 = InStr(fieldCode, plCitStrBeg)

            If n1 > 0 Then
                ' 【审计修复】清除该域结果中已有的超链接，防止重复运行叠加
                Dim hl As Long
                For hl = aField.Result.Hyperlinks.Count To 1 Step -1
                    aField.Result.Hyperlinks(hl).Delete
                Next hl

                n1 = n1 + Len(plCitStrBeg)
                n2 = InStr(Mid(fieldCode, n1, Len(fieldCode) - n1), plCitStrEnd) - 1 + n1
                plain_Cit = Mid$(fieldCode, n1 - 1, n2 - n1 + 2)

                ' 【审计修复】解码所有 \uXXXX Unicode 转义和 \- 转义
                plain_Cit = DecodeUnicodeEscapes(plain_Cit)

                ' ========== 【审计修复】清空数组，防止上一轮迭代的残留数据 ==========
                Erase array_RefTitle
                Erase RefNumber

                ' ========== 解析所有 title ==========
                ' 【审计修复】将转义引号 \" 替换为占位符，防止 JSON 解析被截断
                fieldCode = Replace(fieldCode, "\""", Chr(1))
                i = 0
                ' 【Bug修复】用 FindExactTitleKey 精确匹配 JSON 的 "title" 键，
                ' 排除 "container-title"、"short-title"、"original-title" 等
                Dim searchPos As Long
                searchPos = 1
                Do While searchPos <= Len(fieldCode)
                    Dim foundPos As Long
                    foundPos = FindExactTitleKey(fieldCode, searchPos)
                    If foundPos = 0 Then Exit Do

                    n1 = foundPos + Len("""title"":""")
                    n2 = InStr(Mid(fieldCode, n1, Len(fieldCode) - n1), """,""") - 1 + n1
                    If n2 < n1 Then
                        n2 = InStr(Mid(fieldCode, n1, Len(fieldCode) - n1), "}") - 1 + n1 - 1
                    End If
                    ' 【审计修复】若仍无法定位标题结尾，强制推进防止死循环
                    If n2 < n1 Then n2 = n1
                    ' 【审计修复】还原占位符为真正的引号
                    Dim rawTitle As String
                    rawTitle = DecodeUnicodeEscapes(Replace(Mid(fieldCode, n1, n2 - n1), Chr(1), """"))
                    ' 【Bug修复】剥离 HTML 标签（<i>, <b>, <sup>, <sub> 等），
                    ' 否则带标签的标题在参考文献列表中找不到匹配
                    array_RefTitle(i) = StripHtmlTags(rawTitle)
                    searchPos = n2 + 1
                    i = i + 1
                    If i > 199 Then
                        Exit Do
                    End If
                Loop
                Titles_in_Cit = i

                ' ========== 【Bug修复】解析 RefNumber：拆分复合引文 ==========
                ' 将 "[2,5,7-9]" 拆分为独立数字 "2", "5", "7", "8", "9"
                ' 而非原来的一个整体 "2,5,7-9"
                i = 0
                Do While (InStr(plain_Cit, "]") Or InStr(plain_Cit, "[")) > 0
                    n1 = InStr(plain_Cit, "[")
                    n2 = InStr(plain_Cit, "]")
                    If n1 > 0 And n2 > n1 Then
                        Dim bracketContent As String
                        bracketContent = Mid(plain_Cit, n1 + 1, n2 - (n1 + 1))
                        ' 统一破折号为普通连字符
                        bracketContent = Replace(bracketContent, ChrW(8211), "-")
                        bracketContent = Replace(bracketContent, ChrW(8212), "-")

                        ' 按逗号拆分
                        Dim parts() As String
                        parts = Split(bracketContent, ",")
                        Dim p As Long
                        For p = 0 To UBound(parts)
                            Dim part As String
                            part = Trim(parts(p))
                            If InStr(part, "-") > 0 Then
                                ' 展开范围，如 "7-9" -> 7, 8, 9
                                Dim rangeParts() As String
                                rangeParts = Split(part, "-")
                                If UBound(rangeParts) = 1 Then
                                    Dim rStart As Long, rEnd As Long
                                    On Error Resume Next
                                    rStart = CLng(Trim(rangeParts(0)))
                                    rEnd = CLng(Trim(rangeParts(1)))
                                    On Error GoTo CitationError
                                    If rStart > 0 And rEnd >= rStart And (rEnd - rStart) < 50 Then
                                        Dim r As Long
                                        For r = rStart To rEnd
                                            RefNumber(i) = CStr(r)
                                            i = i + 1
                                            If i > 199 Then Exit For
                                        Next r
                                    Else
                                        ' 无法解析范围，保留原文
                                        RefNumber(i) = part
                                        i = i + 1
                                    End If
                                Else
                                    RefNumber(i) = part
                                    i = i + 1
                                End If
                            Else
                                RefNumber(i) = part
                                i = i + 1
                            End If
                            If i > 199 Then Exit For
                        Next p

                        plain_Cit = Mid(plain_Cit, n2 + 1, Len(plain_Cit) - (n2 + 1) + 1)
                    Else
                        Exit Do
                    End If
                    If i > 199 Then
                        Exit Do
                    End If
                Loop
                Refs_in_Cit = i

                ' ========== 标题映射 ==========
                ' 如果标题数少于引文数，复用最后一个标题填充
                ' 如果标题数多于引文数，清空多余的
                If Titles_in_Cit > Refs_in_Cit And Refs_in_Cit > 0 Then
                    Dim k As Long
                    k = Refs_in_Cit
                    Do While k <= Titles_in_Cit - 1
                        array_RefTitle(k) = ""
                        k = k + 1
                    Loop
                ElseIf Titles_in_Cit < Refs_in_Cit And Titles_in_Cit > 0 Then
                    ' 复合引文展开后，引文数 > 标题数
                    ' 标题按原始顺序对应，多出的引文号无标题则跳过
                    ' 不做额外处理，循环中 title="" 会被 If title <> "" 跳过
                End If

                ' ========== 为每个 Ref 创建超链接 ==========
                For Refs = 0 To Refs_in_Cit - 1 Step 1
                    title = array_RefTitle(Refs)
                    If title <> "" Then
                        array_RefTitle(Refs) = ""
                        titleAnchor = MakeValidBMName(title)

                        ' 【审计修复】防止书签名截断后冲突，冲突时追加数字后缀
                        Dim baseBMName As String
                        Dim bmSuffix As Long
                        baseBMName = titleAnchor
                        bmSuffix = 2
                        Do While usedBookmarks.Exists(titleAnchor)
                            titleAnchor = Left(baseBMName, 40 - Len(CStr(bmSuffix))) & bmSuffix
                            bmSuffix = bmSuffix + 1
                        Loop
                        usedBookmarks(titleAnchor) = True

                        ' --- 在参考文献列表中查找对应条目并打书签 ---
                        ActiveWindow.View.ShowFieldCodes = False
                        Selection.GoTo What:=wdGoToBookmark, Name:="Zotero_Bibliography"
                        ' 【Bug修复】折叠到起点，确保 Find 从参考文献列表开头搜索
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

                        ' 【审计修复】综合验证：Find 结果有效 + 长标题全文匹配 + 在参考文献区域内
                        Dim titleVerified As Boolean
                        titleVerified = Selection.Find.Found

                        ' 标题超 255 字符时，验证匹配位置的完整文本一致
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

                        ' 验证匹配位置确实在参考文献区域内，而非正文
                        If titleVerified And ActiveDocument.Bookmarks.Exists("Zotero_Bibliography") Then
                            Dim biblRange As Range
                            Set biblRange = ActiveDocument.Bookmarks("Zotero_Bibliography").Range
                            If Selection.Start < biblRange.Start Or Selection.End > biblRange.End Then
                                titleVerified = False
                            End If
                        End If

                        If titleVerified Then
                            Selection.MoveStartUntil ("["), Count:=wdBackward
                            Selection.MoveEndUntil (vbCr)
                            lnkcap = Selection.Text
                            lnkcap = Left(lnkcap, 70)

                            Selection.Shrink
                            With ActiveDocument.Bookmarks
                                .Add Range:=Selection.Range, Name:=titleAnchor
                                .DefaultSorting = wdSortByName
                                .ShowHidden = True
                            End With

                            ' 【兜底】确认书签确实创建成功，否则跳过，避免产生断链
                            If Not ActiveDocument.Bookmarks.Exists(titleAnchor) Then GoTo NextRef

                            ' --- 在域结果中搜索 RefNumber，尝试多种破折号变体 ---
                            Dim searchFound As Boolean
                            searchFound = False

                            Dim dashVariants(2) As String
                            dashVariants(0) = RefNumber(Refs)
                            dashVariants(1) = Replace(Replace(RefNumber(Refs), ChrW(8211), "-"), ChrW(8212), "-")
                            dashVariants(2) = Replace(Replace(RefNumber(Refs), "-", ChrW(8211)), ChrW(8212), ChrW(8211))

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

                            ' 兜底：万一所有变体都找不到，给整个引注块加链接
                            If Not searchFound Then
                                ActiveDocument.Hyperlinks.Add Anchor:=aField.Result, _
                                    Address:="", SubAddress:=titleAnchor, ScreenTip:=lnkcap
                            End If

                            ' 去除链接自带的蓝色下划线，保持论文黑字原样
                            aField.Select
                            With Selection.Font
                                 .Underline = wdUnderlineNone
                                 .ColorIndex = wdBlack
                            End With

                        Else
                            ' 【调试】记录未找到标题的引文
                            failedCitations = failedCitations & "[" & RefNumber(Refs) & "] " & Left(title, 30) & "..." & vbCrLf
                        End If
NextRef:
                    End If
                Next Refs
            End If
        GoTo NextCitation
CitationError:
        ' 【兜底】单条引注处理失败：静默跳过，确保不留下蓝色下划线
        On Error Resume Next
        aField.Select
        Selection.Font.Underline = wdUnderlineNone
        Selection.Font.ColorIndex = wdBlack
        On Error GoTo 0
        ' 必须用 Resume 退出错误处理模式，否则下一次迭代的 On Error GoTo 不生效
        Resume NextCitation
NextCitation:
    Next zi
    On Error GoTo ErrorHandler

    ' 恢复原始选择范围并重新开启屏幕刷新
    ActiveWindow.View.ShowFieldCodes = False
    ActiveDocument.Range(nStart, nEnd).Select
    Application.ScreenUpdating = True

    If failedCitations <> "" Then
        MsgBox "引注超链接已生成，但以下引文未能匹配到参考文献列表：" & vbCrLf & vbCrLf & failedCitations, vbExclamation, "部分完成"
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
    ' 【审计修复】收缩连续下划线
    Do While InStr(tempStr, "__") > 0
        tempStr = Replace(tempStr, "__", "_")
    Loop
    MakeValidBMName = Left(tempStr, 40)
End Function

' 【审计修复】通用 Unicode 转义解码：将所有 \uXXXX 替换为对应字符，并处理 \-
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

' 【Bug修复】精确匹配 JSON 中的 "title" 键
' 排除 "container-title"、"short-title"、"original-title"、"title-short" 等
' 原理：找到 "title":" 后，检查其前一个字符是否为 , 或 { 或空白
' 如果前一个字符是 - 或字母，说明匹配到了 xxx-title，跳过继续搜索
Function FindExactTitleKey(ByVal s As String, ByVal startPos As Long) As Long
    Dim target As String
    target = """title"":"""
    Dim pos As Long
    pos = InStr(startPos, s, target)
    Do While pos > 0
        If pos = 1 Then
            ' 在字符串最开头，是精确匹配
            FindExactTitleKey = pos
            Exit Function
        End If
        Dim prevChar As String
        prevChar = Mid(s, pos - 1, 1)
        ' "title" 前面应该是 , 或 { 或空白，表示这是独立的 JSON key
        ' 如果是 - 或字母，则是 container-title、short-title 等的一部分
        If prevChar = "," Or prevChar = "{" Or prevChar = " " Or prevChar = vbTab Or prevChar = vbLf Or prevChar = vbCr Then
            FindExactTitleKey = pos
            Exit Function
        End If
        ' 不是精确匹配，继续搜索下一个
        pos = InStr(pos + 1, s, target)
    Loop
    FindExactTitleKey = 0
End Function

' 【Bug修复】剥离 HTML 标签
' Zotero 域代码 JSON 中的 title 可能含 <i>, <b>, <sup>, <sub> 等标签
' 但参考文献列表中显示的是 Word 格式化后的纯文本，不含这些标签
' 不剥离就会导致 Selection.Find 找不到标题 → 不创建书签 → 超链接失效
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
