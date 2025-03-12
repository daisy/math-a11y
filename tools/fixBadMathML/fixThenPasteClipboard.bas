Attribute VB_Name = "fixBadMathML"
Sub FixThenPasteClipboard()
' This Sub references the Microsoft Forms 2.0 Object Library
' It works well if assigned to ctrl shift alt V

    Dim clipboardData As DataObject
    Dim clipboardText As String

    ' Initialize the DataObject for clipboard interaction
    Set clipboardData = New DataObject
    clipboardData.GetFromClipboard

    ' Get the clipboard content as text
    On Error Resume Next
    clipboardText = clipboardData.GetText
    On Error GoTo 0

    If clipboardText <> "" Then
        ' Check if the string starts with the well formed MathML namespace
        If Left(clipboardText, 48) = "<math xmlns=""http://www.w3.org/1998/Math/MathML""" Then
            ' Replace the clipboard contents with the plain text
            clipboardData.SetText clipboardText
            clipboardData.PutInClipboard
        End If
        ' Check if the string starts with the malformed MathML namespace
        If Left(clipboardText, 43) = "<math xmlns=""//www.w3.org/1998/Math/MathML""" Then
            ' Replace the namespace URL
            clipboardText = Replace(clipboardText, "//www.w3.org/1998/Math/MathML", "http://www.w3.org/1998/Math/MathML")
            ' Replace the clipboard contents with the modified text
            clipboardData.SetText clipboardText
            clipboardData.PutInClipboard
        End If
    End If
    ' Perform the paste operation that we interupted
    Selection.Paste
End Sub
