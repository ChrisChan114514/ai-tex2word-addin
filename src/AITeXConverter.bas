Attribute VB_Name = "AITeXConverter"
Option Explicit

' AI TeX to native Word equations.
' Copyright (c) 2026 Chen Wenxuan.
' Licensed for non-commercial educational and research use only.
' See LICENSE for terms. Commercial use requires prior written permission.

Private Const SelectionShortcutMacro As String = _
    "AITeX_ConvertSelectionMacro"
Private mLastReport As String
Private mConverting As Boolean

Public Sub AutoExec()
    EnsureSelectionShortcut
End Sub

Public Sub AITeX_RibbonOnLoad(ByVal ribbon As Object)
    EnsureSelectionShortcut
End Sub

Private Sub EnsureSelectionShortcut()
    Dim shortcutCode As Long
    Dim existingBinding As KeyBinding
    Dim previousContext As Object

    On Error GoTo Done
    shortcutCode = BuildKeyCode(wdKeyControl, wdKeyAlt, _
                                wdKeyShift, wdKeyE)
    Set existingBinding = Application.FindKey(shortcutCode)

    ' Never replace a Word, user, or third-party shortcut. If this add-in has
    ' already registered the key in the current session, there is no work.
    If Len(existingBinding.Command) > 0 Then Exit Sub

    Set previousContext = Application.CustomizationContext
    Application.CustomizationContext = ThisDocument
    Application.KeyBindings.Add _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:=SelectionShortcutMacro, _
        KeyCode:=shortcutCode

Done:
    On Error Resume Next
    If Not previousContext Is Nothing Then _
        Application.CustomizationContext = previousContext
    ' Register again on every startup instead of modifying the installed file.
    ThisDocument.Saved = True
    On Error GoTo 0
End Sub

Public Sub AITeX_RibbonConvertSelection(ByVal control As Object)
    If Documents.Count = 0 Then Exit Sub
    If Selection.Start = Selection.End Then Exit Sub
    RunConversion Selection.Range.Duplicate
End Sub

Public Sub AITeX_RibbonConvertDocument(ByVal control As Object)
    If Documents.Count = 0 Then Exit Sub
    RunConversion ActiveDocument.Content.Duplicate
End Sub

Public Sub AITeX_RibbonHelp(ByVal control As Object)
End Sub

Public Sub AITeX_ConvertSelectionMacro()
    If Documents.Count = 0 Then Exit Sub
    If Selection.Start <> Selection.End Then _
        RunConversion Selection.Range.Duplicate
End Sub

Public Sub AITeX_ConvertDocumentMacro()
    If Documents.Count = 0 Then Exit Sub
    RunConversion ActiveDocument.Content.Duplicate
End Sub

Public Function AITeX_ConvertDocumentSilent() As Long
    If Documents.Count = 0 Then Exit Function
    AITeX_ConvertDocumentSilent = RunConversion( _
        ActiveDocument.Content.Duplicate)
End Function

Private Function RunConversion(ByVal target As Range) As Long
    Dim count As Long
    Dim oldUpdate As Boolean
    If mConverting Then Exit Function
    mConverting = True
    mLastReport = ""

    If ActiveDocument.ProtectionType <> -1 Then
        mConverting = False
        Exit Function
    End If

    oldUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Done

    count = ConvertMatches(target)

Done:
    If Err.Number <> 0 Then mLastReport = mLastReport & Err.Description
    mConverting = False
    On Error Resume Next
    Application.ScreenUpdating = oldUpdate
    On Error GoTo 0

    RunConversion = count
End Function

Public Function AITeX_GetLastReport() As String
    AITeX_GetLastReport = mLastReport
End Function

Public Function AITeX_GetVersion() As String
    AITeX_GetVersion = "2026.10.31-formatting"
End Function

Public Sub AITeX_ShowLastReport(ByVal control As Object)
    If Len(mLastReport) = 0 Then
        MsgBox "No conversion errors in this session.", vbInformation, "AI TeX"
    Else
        MsgBox Left$(mLastReport, 2000), vbExclamation, "AI TeX"
    End If
End Sub

Private Function ConvertMatches(ByVal target As Range) As Long
    Dim re As Object, matches As Object, oneMatch As Object
    Dim request As Object, response As Object, formula As Object
    Dim result As Object, math As Object, shell As Object, fso As Object
    Dim i As Long, count As Long, failures As Long
    Dim tex As String, displayMode As Boolean, baseAt As Long
    Dim workDir As String, inputPath As String, outputPath As String
    Dim engine As String, command As String, detail As String
    Dim recordStarted As Boolean
    Dim itemRange As Range
    Dim staging As Document

    On Error GoTo Failed
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.MultiLine = True
    re.Pattern = "\\{1,2}\(([\s\S]*?)\\{1,2}\)|\\{1,2}\[([\s\S]*?)\\{1,2}\]"
    Set matches = re.Execute(target.Text)
    If matches.Count = 0 Then Exit Function
    baseAt = target.Start
    engine = ThisDocument.Path & "\AITeXEngine\Convert-TexToOmml.ps1"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(engine) Then _
        Err.Raise vbObjectError + 2200, , "Missing TeX engine. Reinstall the add-in."

    Set request = CreateObject("MSXML2.DOMDocument.6.0")
    request.LoadXML "<request/>"
    For i = 0 To matches.Count - 1
        Set oneMatch = matches.Item(i)
        displayMode = InStr(1, Left$(oneMatch.Value, 3), "[", vbBinaryCompare) > 0
        If displayMode Then
            tex = CStr(oneMatch.SubMatches(1))
        Else
            tex = CStr(oneMatch.SubMatches(0))
        End If
        Set formula = request.createElement("formula")
        formula.setAttribute "id", CStr(i)
        formula.setAttribute "display", CStr(Abs(CInt(displayMode)))
        formula.Text = tex
        request.documentElement.appendChild formula
    Next i

    workDir = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName)
    fso.CreateFolder workDir
    inputPath = workDir & "\request.xml"
    outputPath = workDir & "\response.xml"
    request.Save inputPath
    Set shell = CreateObject("WScript.Shell")
    command = QuoteArg(Environ$("SystemRoot") & _
        "\System32\WindowsPowerShell\v1.0\powershell.exe") & _
        " -NoProfile -NonInteractive -ExecutionPolicy Bypass -File " & _
        QuoteArg(engine) & " -RequestPath " & QuoteArg(inputPath) & _
        " -ResponsePath " & QuoteArg(outputPath)
    Application.StatusBar = "AI TeX: parsing " & matches.Count & " equations..."
    If shell.Run(command, 0, True) <> 0 Then _
        Err.Raise vbObjectError + 2201, , "The TeX engine could not start."
    Set response = CreateObject("MSXML2.DOMDocument.6.0")
    response.async = False
    If Not response.Load(outputPath) Then _
        Err.Raise vbObjectError + 2202, , "The TeX engine returned no valid response."
    If Not IsNull(response.documentElement.getAttribute("error")) Then _
        Err.Raise vbObjectError + 2203, , response.documentElement.getAttribute("error")
    response.setProperty "SelectionNamespaces", _
        "xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'"

    ' Import XML only into a private document, before recording target edits.
    ' Copying an equation range excludes paragraph marks and document styles.
    Set staging = PrepareNativeEquations(response)
    target.Document.Activate
    Application.UndoRecord.StartCustomRecord "AI TeX"
    recordStarted = True

    ' Resolve every formula before editing, then insert in reverse order.
    For i = matches.Count - 1 To 0 Step -1
        Set oneMatch = matches.Item(i)
        Set result = response.selectSingleNode("/results/result[@id='" & CStr(i) & "']")
        detail = "No parser result."
        Set math = Nothing
        If Not result Is Nothing Then
            If Not IsNull(result.getAttribute("error")) Then
                detail = result.getAttribute("error")
            Else
                Set math = result.selectSingleNode("m:oMath")
            End If
        End If
        If Not math Is Nothing Then
            displayMode = InStr(1, Left$(oneMatch.Value, 3), "[", vbBinaryCompare) > 0
            Set itemRange = target.Document.Range(baseAt + oneMatch.FirstIndex, _
                baseAt + oneMatch.FirstIndex + oneMatch.Length)
            If InsertNativeEquation(itemRange, staging.OMaths.Item( _
                CLng(result.getAttribute("stagingIndex"))).Range, displayMode, detail) Then
                count = count + 1
            Else
                failures = failures + 1
            End If
        Else
            failures = failures + 1
        End If
        If Len(detail) > 0 Then mLastReport = mLastReport & _
            "Formula " & CStr(i + 1) & ": " & detail & vbCrLf & _
            oneMatch.Value & vbCrLf & vbCrLf
    Next i
    GoTo Done
Failed:
    mLastReport = mLastReport & Err.Description & vbCrLf
    If Not matches Is Nothing Then failures = matches.Count - count
Done:
    On Error Resume Next
    If recordStarted Then Application.UndoRecord.EndCustomRecord
    target.Document.Activate
    If Not staging Is Nothing Then staging.Close wdDoNotSaveChanges
    ' Only the unique directory created above belongs to this conversion.
    If Len(workDir) > 0 Then
        If StrComp(fso.GetParentFolderName(fso.GetAbsolutePathName(workDir)), _
            fso.GetAbsolutePathName(CStr(fso.GetSpecialFolder(2))), vbTextCompare) = 0 Then _
            fso.DeleteFolder workDir, True
    End If
    Application.StatusBar = "AI TeX: " & count & " converted; " & _
        failures & " preserved. See conversion errors for details."
    ConvertMatches = count
End Function

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr$(34) & value & Chr$(34)
End Function

Private Function PrepareNativeEquations(ByVal response As Object) As Document
    Dim staging As Document
    Dim package As Object, body As Object, para As Object, result As Object, math As Object
    Dim count As Long, errorNumber As Long, errorText As String

    On Error GoTo Failed
    If response.selectNodes("/results/result/m:oMath").Length = 0 Then Exit Function
    Set staging = Documents.Add(Visible:=False)
    Set package = CreateObject("MSXML2.DOMDocument.6.0")
    If Not package.LoadXML(staging.Content.WordOpenXML) Then _
        Err.Raise vbObjectError + 2205, , "Cannot prepare the equation document."
    package.setProperty "SelectionNamespaces", _
        "xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'"
    Set body = package.selectSingleNode("//w:body")
    Do While body.hasChildNodes
        body.removeChild body.firstChild
    Loop
    For Each result In response.selectNodes("/results/result")
        Set math = result.selectSingleNode("m:oMath")
        If Not math Is Nothing Then
            Set para = package.createNode(1, "w:p", _
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            para.appendChild package.importNode(math, True)
            body.appendChild para
            count = count + 1
            result.setAttribute "stagingIndex", CStr(count)
        End If
    Next result
    staging.Content.InsertXML package.XML
    If staging.OMaths.Count <> count Then _
        Err.Raise vbObjectError + 2206, , "Cannot prepare all native equations."
    Set PrepareNativeEquations = staging
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If Not staging Is Nothing Then staging.Close wdDoNotSaveChanges
    On Error GoTo 0
    Err.Raise errorNumber, , errorText
End Function

Private Function InsertNativeEquation(ByVal source As Range, ByVal mathSource As Range, _
                                      ByVal displayMode As Boolean, _
                                      ByRef detail As String) As Boolean
    Dim original As String, startAt As Long, doc As Document
    Dim equation As OMath, inserted As Range
    original = source.Text
    startAt = source.Start
    Set doc = source.Document
    On Error GoTo Failed
    ' OMath.Range contains no paragraph mark. Importing that range preserves
    ' the target paragraph's style, spacing, indents, and surrounding runs.
    source.FormattedText = mathSource.FormattedText
    Set inserted = doc.Range(startAt, source.End)
    If inserted.OMaths.Count <> 1 Then _
        Err.Raise vbObjectError + 2204, , "Word did not insert one native equation."
    Set equation = inserted.OMaths.Item(1)
    If displayMode Then
        If equation.Type <> wdOMathDisplay Then equation.Type = wdOMathDisplay
    Else
        ' Setting Type again can invoke Word's paragraph-layout conversion.
        ' Inline math already imported into prose must stay in that text flow.
        If equation.Type <> wdOMathInline Then equation.Type = wdOMathInline
    End If
    detail = ""
    InsertNativeEquation = True
    Exit Function
Failed:
    detail = Err.Description
    On Error Resume Next
    source.Text = original
End Function
