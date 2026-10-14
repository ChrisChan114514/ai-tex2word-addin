Attribute VB_Name = "AITeXConverter"
Option Explicit

' AI TeX to native Word equations.
' Copyright (c) 2026 Chen Wenxuan.
' Licensed for non-commercial educational and research use only.
' See LICENSE for terms. Commercial use requires prior written permission.

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

    If ActiveDocument.ProtectionType <> -1 Then Exit Function

    oldUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Done

    Application.UndoRecord.StartCustomRecord "AI"
    count = ConvertMatches(target)

Done:
    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = oldUpdate
    On Error GoTo 0

    RunConversion = count
End Function

Private Function ConvertMatches(ByVal target As Range) As Long
    Dim re As Object
    Dim allMatches As Object
    Dim oneMatch As Object
    Dim i As Long
    Dim displayMode As Boolean
    Dim tex As String
    Dim itemRange As Range
    Dim count As Long

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.MultiLine = True
    ' Accept normal TeX delimiters and their JSON/Markdown-escaped form:
    ' \(...\), \[...\], \\(...\\), and \\[...\\].
    re.Pattern = "\\{1,2}\(([\s\S]*?)\\{1,2}\)|\\{1,2}\[([\s\S]*?)\\{1,2}\]"
    Set allMatches = re.Execute(target.Text)

    For i = allMatches.Count - 1 To 0 Step -1
        Set oneMatch = allMatches.Item(i)
        displayMode = (InStr(1, Left$(oneMatch.Value, 3), "[", _
                             vbBinaryCompare) > 0)
        If displayMode Then
            tex = CStr(oneMatch.SubMatches(1))
        Else
            tex = CStr(oneMatch.SubMatches(0))
        End If

        Set itemRange = target.Document.Range( _
            target.Start + oneMatch.FirstIndex, _
            target.Start + oneMatch.FirstIndex + oneMatch.Length)

        If ConvertOne(itemRange, tex, displayMode) Then count = count + 1
    Next i

    ConvertMatches = count
End Function

Private Function ConvertOne(ByVal sourceRange As Range, _
                            ByVal tex As String, _
                            ByVal displayMode As Boolean) As Boolean
    Dim doc As Document
    Dim original As String
    Dim startAt As Long
    Dim mathRange As Range
    Dim resultRange As Range
    Dim equation As OMath

    If Len(Trim$(tex)) = 0 Then Exit Function

    Set doc = sourceRange.Document
    original = sourceRange.Text
    startAt = sourceRange.Start
    tex = NormalizeTex(tex)

    On Error GoTo Failed
    sourceRange.Text = tex
    Set mathRange = doc.Range(startAt, startAt + Len(tex))
    Set resultRange = doc.OMaths.Add(mathRange)
    Set equation = resultRange.OMaths.Item(1)
    equation.BuildUp
    If displayMode Then
        equation.Type = 0
    Else
        equation.Type = 1
    End If
    ConvertOne = True
    Exit Function

Failed:
    On Error Resume Next
    If Not resultRange Is Nothing Then
        doc.Range(startAt, resultRange.End).Text = original
    Else
        doc.Range(startAt, startAt + Len(tex)).Text = original
    End If
    On Error GoTo 0
End Function

Private Function NormalizeTex(ByVal tex As String) As String
    Dim scriptCommand As Object
    Dim operatorName As Object

    tex = Trim$(tex)
    tex = Replace(tex, "\displaystyle", "")
    tex = Replace(tex, "\dfrac", "\frac")
    tex = Replace(tex, "\tfrac", "\frac")

    ' AI and Markdown output sometimes escapes underscores inside math as \_.
    ' Word would render that as a literal underscore instead of a subscript.
    tex = Replace(tex, "\_", "_")

    ' Word's linear equation parser needs a command used as a script to be
    ' grouped. For example, C_\pi loses the final letter, while C_{\pi}
    ' builds correctly. Apply the same repair to superscripts and to any
    ' alphabetic command (Greek letters, \infty, and similar commands).
    Set scriptCommand = CreateObject("VBScript.RegExp")
    scriptCommand.Global = True
    scriptCommand.Pattern = "([_^])\\([A-Za-z]+)"
    tex = scriptCommand.Replace(tex, "$1{\$2}")

    ' \operatorname is common in AI-generated TeX but is not understood by
    ' Word's linear parser. The equivalent upright text style is supported.
    Set operatorName = CreateObject("VBScript.RegExp")
    operatorName.Global = True
    operatorName.Pattern = "\\operatorname\*?\{([^{}]+)\}"
    tex = operatorName.Replace(tex, "\mathrm{$1}")

    NormalizeTex = Trim$(tex)
End Function
