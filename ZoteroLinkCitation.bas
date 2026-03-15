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
    Dim truncatedCount As Long
    truncatedCount = 0
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

    ' 倒序遍历：新插入的 Field 索引在后方，不影响前方未处理的元素
    Dim zi As Long
    For zi = zotFieldCount To 1 Step -1
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
                Do While InStr(fieldCode, """title"":""") > 0
                    n1 = InStr(fieldCode, """title"":""") + Len("""title"":""")
                    n2 = InStr(Mid(fieldCode, n1, Len(fieldCode) - n1), """,""") - 1 + n1
                    If n2 < n1 Then
                        n2 = InStr(Mid(fieldCode, n1, Len(fieldCode) - n1), "}") - 1 + n1 - 1
                    End If
                    ' 【审计修复】若仍无法定位标题结尾，强制推进防止死循环
                    If n2 < n1 Then n2 = n1
                    ' 【审计修复】还原占位符为真正的引号
                    array_RefTitle(i) = DecodeUnicodeEscapes(Replace(Mid(fieldCode, n1, n2 - n1), Chr(1), """"))
                    fieldCode = Mid(fieldCode, n2 + 1, Len(fieldCode) - n2 - 1)
                    i = i + 1
                    If i > 199 Then
                        truncatedCount = truncatedCount + 1
                        Exit Do
                    End If
                Loop
                Titles_in_Cit = i
                
                ' ========== 解析所有 RefNumber ==========
                i = 0
                Do While (InStr(plain_Cit, "]") Or InStr(plain_Cit, "[")) > 0
                    n1 = InStr(plain_Cit, "[")
                    n2 = InStr(plain_Cit, "]")
                    If n1 > 0 And n2 > n1 Then
                        RefNumber(i) = Mid(plain_Cit, n1 + 1, n2 - (n1 + 1))
                        plain_Cit = Mid(plain_Cit, n2 + 1, Len(plain_Cit) - (n2 + 1) + 1)
                        i = i + 1
                    Else
                        Exit Do
                    End If
                    If i > 199 Then
                        truncatedCount = truncatedCount + 1
                        Exit Do
                    End If
                Loop
                Refs_in_Cit = i
                
                ' ========== 【修复1】标题映射：保留前 N 个 title，清空多余的 ==========
                If Titles_in_Cit > Refs_in_Cit And Refs_in_Cit > 0 Then
                    i = Refs_in_Cit
                    Do While i <= Titles_in_Cit - 1
                        array_RefTitle(i) = ""
                        i = i + 1
                    Loop
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
                            
                            ' --- 【修复2】在域结果中搜索 RefNumber，尝试多种破折号变体 ---
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
                            
                        End If
                    End If
                Next Refs
            End If
    Next zi

    ' 恢复原始选择范围并重新开启屏幕刷新
    ActiveWindow.View.ShowFieldCodes = False
    ActiveDocument.Range(nStart, nEnd).Select
    Application.ScreenUpdating = True
    
    ' 【审计修复】如有数组截断，提示用户部分引注未处理
    If truncatedCount > 0 Then
        MsgBox "处理完成，但有 " & truncatedCount & " 处引注因超过 200 条上限被截断，部分链接可能缺失。", vbExclamation, "完成（有警告）"
    Else
        MsgBox "完美搞定！包含合并格式的引注超链接已全部生成完毕！", vbInformation, "完成"
    End If

    Exit Sub

ErrorHandler:
    ActiveWindow.View.ShowFieldCodes = False
    Application.ScreenUpdating = True
    MsgBox "运行出错: " & Err.Description & "（错误号 " & Err.Number & "）", vbCritical, "错误"

End Sub

Function MakeValidBMName(strIn As String)
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
    ' 【审计修复】收缩连续下划线（原代码误用双空格，实际不可能出现空格）
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
