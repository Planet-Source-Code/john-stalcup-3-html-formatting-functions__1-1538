<div align="center">

## 3 HTML Formatting Functions


</div>

### Description

This code provides 3 convenient ways of formatting html strings.
 
### More Info
 
A string of html.

A formatted string of html


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Stalcup](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-stalcup.md)
**Level**          |Unknown
**User Rating**    |3.0 (9 globes from 3 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-stalcup-3-html-formatting-functions__1-1538/archive/master.zip)





### Source Code

```
Option Explicit
Type tag
  text As String
  start As Double
  length As Double
End Type
'*********************************************************************
Public Function SimpleFormat(target As String) As String
SimpleFormat = ReplaceSubString(CompactFormat(target), "><", ">" & vbCrLf & "<")
End Function
'*********************************************************************
Public Function CompactFormat(target As String) As String
Dim a As String
a = ReplaceSubString(target, vbCrLf, "")
a = ReplaceSubString(a, Chr(9), " ")
a = ReplaceSubString(a, "   ", " ")
a = ReplaceSubString(a, "  ", " ")
a = ReplaceSubString(a, "  ", " ")
a = ReplaceSubString(a, " ", " ")
a = Clean(a)
CompactFormat = a
End Function
'*********************************************************************
Public Function HierarchalFormat(target As String) As String
  target = ReplaceSubString(target, vbCrLf, "")
  target = ReplaceSubString(target, vbTab, "")
  target = Eformat(target)
  HierarchalFormat = Clean(target)
End Function
'*********************************************************************
'this lines denotes separation from public access and inner workings
'*********************************************************************
Private Function Clean(targ As String) As String
targ = ReplaceSubString(targ, " >", ">")
targ = ReplaceSubString(targ, "< ", "<")
targ = ReplaceSubString(targ, "> <", "><")
Clean = targ
End Function
Public Function ReplaceSubString(str As String, ByVal substr As String, ByVal newsubstr As String)
Dim pos As Double
Dim startPos As Double
Dim new_str As String
  startPos = 1
  pos = InStr(str, substr)
  Do While pos > 0
    new_str = new_str & Mid$(str, startPos, pos - startPos) & newsubstr
    startPos = pos + Len(substr)
    pos = InStr(startPos, str, substr)
  Loop
  new_str = new_str & Mid$(str, startPos)
  ReplaceSubString = new_str
End Function
Private Function Eformat(str As String) As String
  On Error Resume Next
  Dim startPos As Double
  Dim endPos As Double
  Dim indentationLevel As Double
  Dim new_str As String
  indentationLevel = 0
  startPos = 0
  endPos = 0
  If (Mid$(str, 1, 1) <> "<") Then
    Dim tempEnd As Double
    tempEnd = InStr(1, str, "<")
    If tempEnd = 0 Then
      tempEnd = Len(str)
    End If
    new_str = Mid$(str, 1, tempEnd)
  End If
  Do
    DoEvents
    If InStr(startPos + 1, str, "</") <> 0 And InStr(startPos + 1, str, "</") <= InStr(startPos + 1, str, "<") Then
      startPos = InStr(startPos + 1, str, "</")
      endPos = InStr(startPos + 1, str, "<")
      If endPos = 0 Then
        endPos = Len(str) + 1
      End If
      indentationLevel = indentationLevel - 1
      new_str = new_str & vbCrLf & String(indentationLevel, vbTab) & Mid$(str, startPos, endPos - startPos)
    Else
      startPos = InStr(startPos + 1, str, "<")
      endPos = InStr(startPos + 1, str, "<")
      If endPos = 0 Then
        endPos = Len(str) + 1
      End If
      new_str = new_str & vbCrLf & String(indentationLevel, vbTab) & Mid$(str, startPos, endPos - startPos)
      Dim tagName As String
      tagName = LCase(returnNameOfTag(returnNextTag(str, startPos)))
      If tagName <> "br" And tagName <> "hr" And tagName <> "img" And tagName <> "meta" And tagName <> "applet" And tagName <> "p" And tagName <> "!--" And tagName <> "input" And tagName <> "!doctype" And tagName <> "area" Then
        indentationLevel = indentationLevel + 1
      End If
    End If
  Loop While startPos > 0
  Eformat = new_str
End Function
Public Function returnNextTag(ByRef str As String, ByVal start As Double) As tag
  On Error Resume Next
  Dim endPos As Double
  start = InStr(start + 1, str, "<")
  endPos = InStr(start + 1, str, ">")
  returnNextTag.text = Mid$(str, start, endPos - start + 1)
  returnNextTag.start = start
  returnNextTag.length = endPos - start
End Function
Public Function returnNameOfTag(ByRef str As tag) As String
  On Error Resume Next
  Dim endPos As Double
  Dim start As Double
  start = 2
  endPos = InStr(1, str.text, " ")
  If Mid$(str.text, 2, 3) = "!--" Then
    endPos = 5
  ElseIf endPos = 0 Then
    endPos = InStr(1, str.text, ">")
  End If
  returnNameOfTag = Mid$(str.text, start, endPos - start)
End Function
```

