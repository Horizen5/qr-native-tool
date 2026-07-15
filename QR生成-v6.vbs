Sub QR()
    Const NORMAL_CHUNK As Long = 1000
    Dim f$, data$, t$, fileName$, fileNameB64$, checksum$
    Dim i&, n&, pos&, firstChunkLen&
    Dim stm, x, nd, fnStream, fnNode

    With Application.FileDialog(3)
        If .Show = 0 Then Exit Sub
        f = .SelectedItems(1)
    End With

    fileName = Mid$(f, InStrRev(f, "\") + 1)

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open: stm.LoadFromFile f

    Set x = CreateObject("MSXML2.DOMDocument")
    Set nd = x.createElement("x")
    nd.DataType = "bin.base64"
    nd.nodeTypedValue = stm.Read
    data = Replace(Replace(nd.Text, vbCr, ""), vbLf, "")
    stm.Close

    ' 文件名用 UTF-8 Base64 编码，避免中文/特殊字符进二维码字段
    Set fnStream = CreateObject("ADODB.Stream")
    fnStream.Type = 2
    fnStream.Mode = 3
    fnStream.Charset = "UTF-8"
    fnStream.Open
    fnStream.WriteText fileName
    fnStream.Position = 0
    fnStream.Type = 1
    fnStream.Position = 0

    Set fnNode = x.createElement("f")
    fnNode.DataType = "bin.base64"
    fnNode.nodeTypedValue = fnStream.Read
    fileNameB64 = Replace(Replace(fnNode.Text, vbCr, ""), vbLf, "")
    fnStream.Close

    ' 校验码：文件名Base64前6位 + Base64总长度（纯ASCII，避免中文导致Word字段报错）
    checksum = Left$(fileNameB64, 6) & Len(data)

    ' 第一张除了文件名外，剩余空间给数据；其他张固定 1000，保持和原版接近
    ' Base64 解码要求总长度是 4 的倍数，所以 firstChunkLen 也要对齐 4
    firstChunkLen = ((NORMAL_CHUNK - Len(fileNameB64) - 1) \ 4) * 4
    If firstChunkLen < 200 Then
        MsgBox "文件名太长，请缩短到 100 个字符以内再试。", vbCritical
        Exit Sub
    End If

    If Len(data) <= firstChunkLen Then
        n = 1
    Else
        n = 1 + (Len(data) - firstChunkLen + NORMAL_CHUNK - 1) \ NORMAL_CHUNK
    End If

    Documents.Add
    pos = 1

    For i = 1 To n
        If i = 1 Then
            ' 第一页：第5字段放Base64文件名，第6字段放数据
            t = "TZ|" & checksum & "|1|" & n & "|" & fileNameB64 & "|" & Mid$(data, pos, firstChunkLen)
            pos = pos + firstChunkLen
        Else
            ' 其他页：第5字段空，第6字段放数据
            t = "TZ|" & checksum & "|" & i & "|" & n & "||" & Mid$(data, pos, NORMAL_CHUNK)
            pos = pos + NORMAL_CHUNK
        End If

        Selection.TypeText i & "/" & n & vbCr

        On Error Resume Next
        ActiveDocument.Fields.Add Selection.Range, , _
            "DISPLAYBARCODE """ & t & """ QR \q 3"
        If Err.Number <> 0 Then
            MsgBox "第 " & i & " 页二维码生成失败。" & vbCrLf & _
                   "字段长度：" & Len(t) & vbCrLf & _
                   "前 80 字符：" & Left(t, 80) & vbCrLf & _
                   "错误信息：" & Err.Description, vbCritical
            Exit Sub
        End If
        On Error GoTo 0

        Selection.EndKey 6
        If i < n Then Selection.InsertBreak 7
    Next i

    ActiveDocument.Fields.Update
    MsgBox "完成，共 " & n & " 张，校验码：" & checksum
End Sub
