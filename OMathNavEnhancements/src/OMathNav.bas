Attribute VB_Name = "OMathNav"
' ==========================================================================
' Module:     OMathNav
' Purpose:    Supports non-visual navigation of mathzones (OMath objects) in MS Word
' Author:     Brian Richwine
' Date:       November 5, 2024
' ==========================================================================
' Dependencies:
' ==========================================================================
' Revision History:
'   - v1.0 (2024-11-05): Initial version
' ==========================================================================
' Notes:
'   - Moving from just before a mathzone to inside a mathzone (equation
'     editor appearing) but before the expression does not raise any
'     events nor does it change the Selection.Start
'
'   - The same is true for moving into the equation editor from the end
'
'   - The above is why events cannot be used to detect enter/exit
'
'   - To determine if really in the equation editor cannot depend on
'     Selection.OMaths.Count so need to also check the current font to
'     be a math font
'
'   - To move into a math expression and invoke the editor, cannot simply
'     set Selection.SetRange to start/end of the OMath or the Equation Editor
'     will not appear.
'        * For a display equation must set Selection Start and End to the
'          opposite end first and then back to the desired end.
'        * For inline equations also need to use MoveRight or MoveLeft one
'          wdCharacter.
' ==========================================================================
Option Explicit

#If VBA7 Then
    ' This version works on 64-bit systems
    Declare PtrSafe Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
#Else
    ' This version is for 32-bit systems
    Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
#End If

Private Type AutoCorrectMap
    UnicodeChar As String
    Keyword As String
End Type

Dim blnAnnounceMathZones As Boolean ' Turn announcement on/off
Dim blnMathZoneDetected As Boolean ' Track detection state
Dim blnWavFilesWork As Boolean ' Keep track because On Error is slow
Private lngCallCount As Long  ' Global counter we can watch
Private dtmNextScheduledTime As Date  '

' Initialize the globals
Public Sub OnDocumentOpen()
    'Debug.Print "OnDocumentOpen"
    blnAnnounceMathZones = False
    blnMathZoneDetected = False
    blnWavFilesWork = True
    lngCallCount = 0
    SetShortcutKeys
End Sub

Private Sub Document_Open()
    OnDocumentOpen
End Sub

Sub SetShortcutKeys()
    'Debug.Print "Setting shortcut keys"
    ' Set customization context to the current document or template
    CustomizationContext = ThisDocument.AttachedTemplate
    
    ' Add key bindings for math navigation and functions
    AddKeyBinding BuildKeyCode(wdKeyAlt, 219), "MoveToOMathStart"     ' Alt + [
    AddKeyBinding BuildKeyCode(wdKeyAlt, 221), "MoveToOMathEnd"       ' Alt + ]
    AddKeyBinding BuildKeyCode(wdKeyAlt, wdKeyShift, 219), "MoveToPreviousOMath" ' Alt + Shift + [  (for {)
    AddKeyBinding BuildKeyCode(wdKeyAlt, wdKeyShift, 221), "MoveToNextOMath"     ' Alt + Shift + ]  (for })
    AddKeyBinding BuildKeyCode(wdKeyAlt, 69), "SelectEntireOMath"     ' Alt + E
    AddKeyBinding BuildKeyCode(wdKeyAlt, 220), "OMathStats"           ' Alt + \
    AddKeyBinding BuildKeyCode(wdKeyAlt, wdKeyShift, 220), "TurnOnEnterExitOMathAnnouncement" ' Alt + Shift + \  (for |)
    AddKeyBinding BuildKeyCode(wdKeyControl, wdKeyAlt, 219), "SelectToOMathStart" ' Ctrl + Alt + [  (for {)
    AddKeyBinding BuildKeyCode(wdKeyControl, wdKeyAlt, 221), "SelectToOMathEnd"     'Ctrl + Alt + ]  (for })
    AddKeyBinding BuildKeyCode(wdKeyControl, wdKeyAlt, 220), "EditMathWithForm"     'Ctrl + Alt + \  (for })
    
End Sub

Sub AddKeyBinding(fullKeyCode As Long, macroName As String)
    ' Remove any existing binding for this key code
    On Error Resume Next
    KeyBindings(fullKeyCode).Clear
    On Error GoTo 0
    
    ' Add the new key binding
    KeyBindings.Add keyCode:=fullKeyCode, _
                    KeyCategory:=wdKeyCategoryMacro, _
                    Command:=macroName
End Sub


Public Sub AnnounceMathZone()
    Static lastInMath As Boolean  ' Add state tracking for math zone
    
    Dim blnMathZoneHadBeenDetected As Boolean
    Dim blnInMathFont As Boolean
    
    DoEvents ' This code can hog the CPU due to the frequency at which the sub is called
        
    ' Check for only one OMaths and math font
    If Selection.OMaths.Count > 0 Then
        ' Verify by getting math font state
        blnMathZoneDetected = (InStr(1, Selection.Font.Name, "Math", vbTextCompare) > 0)
        DoEvents
    Else
        blnMathZoneDetected = False
    End If
    
    ' Only process state changes
    If blnMathZoneDetected <> lastInMath Then
        ' Compare current and previous state to detect entry or exit from math zone
        If Not lastInMath And blnMathZoneDetected Then
            Debug.Print "Play Tick"
            playTickSound
        ElseIf lastInMath And Not blnMathZoneDetected Then
            Debug.Print "Play Tock"
            playTockSound
        End If
        
        ' Update cached states
        lastInMath = blnMathZoneDetected
    End If
    
    DoEvents
    
    ' Schedule next check if still enabled
    If blnAnnounceMathZones Then
'       dtmNextScheduledTime = Now() + 0.0000056
        dtmNextScheduledTime = Now() + 0.0000056
        Application.OnTime When:=dtmNextScheduledTime, Name:="OMathNav.AnnounceMathZone"
    End If
End Sub

Public Sub TurnOnEnterExitOMathAnnouncement()
    blnAnnounceMathZones = True
    blnWavFilesWork = True
    dtmNextScheduledTime = Now() + 0.000002  ' Schedule almost immediately
    Application.OnTime When:=dtmNextScheduledTime, Name:="OMathNav.AnnounceMathZone"
End Sub

Public Sub TurnOffEnterExitOMathAnnouncement()
    blnAnnounceMathZones = False
End Sub


Public Sub SelectToOMathStart()
    Dim oMath As oMath
    
    ' Check if the selection is in one OMath object
    If CheckIfSelectionIsInOMath() = True Then
        ' Get the first OMath object in the selection
        Set oMath = Selection.OMaths(1)
        
        ' Select from the current position to the start of the OMath object
        Selection.SetRange Start:=oMath.Range.Start, End:=Selection.End
    Else
        Beep
    End If
End Sub

Public Sub SelectToOMathEnd()
    Dim oMath As oMath
    
    ' Check if the selection is in one OMath object
    If CheckIfSelectionIsInOMath() = True Then
        ' Get the first OMath object in the selection
        Set oMath = Selection.OMaths(1)
        
        ' Select from the current position to the end of the OMath object
        Selection.SetRange Start:=Selection.Start, End:=oMath.Range.End
    Else
        Beep
    End If
End Sub

Public Sub MoveToOMathStart()
    Dim oMath As oMath
    
    ' Check if the selection is in one OMath object
    If CheckIfSelectionIsInOMath() = True Then
        ' Get the first OMath object in the selection
        Set oMath = Selection.OMaths(1)
        
        ' Set both Start and End position to the start of the OMath object
        Selection.SetRange Start:=oMath.Range.Start, End:=oMath.Range.Start
        If CheckIfSelectionIsInOMath() = False Then
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
        End If
    Else
        Beep
    End If
End Sub

Public Sub MoveToOMathEnd()
    Dim oMath As oMath
    
    ' Check if the selection is in one OMath object
    If CheckIfSelectionIsInOMath() = True Then
        ' Get the first OMath object in the selection
        Set oMath = Selection.OMaths(1)
        
        ' Set both the Start and End position to the end of the OMath object
        Selection.SetRange Start:=oMath.Range.End, End:=oMath.Range.End
        If CheckIfSelectionIsInOMath() = False Then
            Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
        End If
    Else
        Beep
    End If
End Sub

Public Sub SelectEntireOMath()
    Dim oMath As oMath
    
    ' Check if the selection is in one OMath object
    If CheckIfSelectionIsInOMath() = True Then
        ' Get the first OMath object in the selection
        Set oMath = Selection.OMaths(1)
        
        ' Expand the selection to include the entire OMath
        oMath.Range.Select
    Else
        Beep
    End If
End Sub

Public Sub MoveToPreviousOMath()
    Dim lngCurrentPosition As Long
    Dim lngLow As Long
    Dim lngHigh As Long
    Dim lngMid As Long
    Dim oMath As oMath
    Dim lngPreviousMathStart As Long
    Dim lngPreviousMathEnd As Long
    
    ' Get the current position of the selection start
    lngCurrentPosition = Selection.Start
    
    If Selection.OMaths.Count = 1 And CheckIfSelectionIsInOMath() = False Then
        Set oMath = Selection.OMaths(1)
        If lngCurrentPosition = oMath.Range.End Then
            ' Move to start then back to end so Equation Editor appears
            Selection.SetRange Start:=oMath.Range.Start, End:=oMath.Range.Start
            Selection.SetRange Start:=oMath.Range.End, End:=oMath.Range.End
            If CheckIfSelectionIsInOMath() = False Then
                Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
            End If
            Exit Sub
        End If
    End If
    
    ' Initialize binary search bounds
    lngLow = 1
    lngHigh = ActiveDocument.OMaths.Count
    lngPreviousMathStart = -1 ' Initialize as no previous expression found
    
    ' Perform binary search
    While lngLow <= lngHigh And lngHigh > 0
        lngMid = (lngLow + lngHigh) \ 2 ' Find the midpoint
        
        ' Get the midpoint OMath object
        Set oMath = ActiveDocument.OMaths(lngMid)
        
        ' Check if the math expression is before the current selection
        If oMath.Range.End < lngCurrentPosition Then
            ' This could be a potential previous math expression
            lngPreviousMathStart = oMath.Range.Start
            lngPreviousMathEnd = oMath.Range.End
            ' Narrow the search to higher indices
            lngLow = lngMid + 1
        Else
            ' Narrow the search to lower indices
            lngHigh = lngMid - 1
        End If
    Wend
    
    ' Check if a previous math expression was found
    If lngPreviousMathStart <> -1 Then
        ' Move the selection to the start of the next math expression; works for Display equations
        Selection.SetRange Start:=lngPreviousMathEnd, End:=lngPreviousMathEnd
        Selection.SetRange Start:=lngPreviousMathStart, End:=lngPreviousMathStart
        ' Necessary for Inline equations
        If CheckIfSelectionIsInOMath() = False Then
            Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
        End If
    Else
        Beep
    End If
End Sub

Public Sub MoveToNextOMath()
    Dim lngCurrentPosition As Long
    Dim lngLow As Long
    Dim lngHigh As Long
    Dim lngMid As Long
    Dim oMath As oMath
    Dim lngNextMathStart As Long
    Dim lngNextMathEnd As Long
    
    
    ' Get the current position of the selection start
    lngCurrentPosition = Selection.Start
    
    If Selection.OMaths.Count = 1 And CheckIfSelectionIsInOMath() = False Then
        Set oMath = Selection.OMaths(1)
        If lngCurrentPosition = oMath.Range.Start Then
            ' Move to end then back to start so Equation Editor appears
            Selection.SetRange Start:=oMath.Range.End, End:=oMath.Range.End
            Selection.SetRange Start:=oMath.Range.Start, End:=oMath.Range.Start
            ' Necessary for Inline equations
            If CheckIfSelectionIsInOMath() = False Then
                Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
            End If
            Exit Sub
        End If
    End If
    
    ' Initialize binary search bounds
    lngLow = 1
    lngHigh = ActiveDocument.OMaths.Count
    lngNextMathStart = -1 ' Initialize as no next expression found
    
    ' Perform binary search
    While lngLow <= lngHigh
        lngMid = (lngLow + lngHigh) \ 2 ' Find the midpoint
        
        ' Get the midpoint OMath object
        Set oMath = ActiveDocument.OMaths(lngMid)
        
        ' Check if the math expression is after the current selection
        If oMath.Range.Start > lngCurrentPosition Then
            ' This could be a potential next math expression
            lngNextMathStart = oMath.Range.Start
            lngNextMathEnd = oMath.Range.End
            ' Narrow the search to lower indices
            lngHigh = lngMid - 1
        Else
            ' Narrow the search to higher indices
            lngLow = lngMid + 1
        End If
    Wend
    
    ' Check if a next math expression was found
    If lngNextMathStart <> -1 Then
        ' Move the selection to the start of the next math expression; works for Display equations
        Selection.SetRange Start:=lngNextMathEnd, End:=lngNextMathEnd
        Selection.SetRange Start:=lngNextMathStart, End:=lngNextMathStart
        ' Necessary for Inline equations
        If CheckIfSelectionIsInOMath() = False Then
         Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
        End If
    Else
        Beep
    End If
End Sub

Private Function CheckIfSelectionIsInOMath() As Boolean
    Dim blnInMathFont As Boolean
    
    DoEvents
    ' Check if the font at the current selection contains "Math"
    blnInMathFont = (InStr(1, Selection.Font.Name, "Math", vbTextCompare) > 0)
    CheckIfSelectionIsInOMath = blnInMathFont And Selection.OMaths.Count = 1
End Function

Sub OMathStats()
    Dim oMathCount As Integer
    Dim i As Integer
    Dim currentOMathIndex As Integer
    Dim isOnOMath As Boolean
    Dim selStart As Long
    Dim selEnd As Long
    Dim message As String
    Dim beforeFirst As Boolean
    Dim afterLast As Boolean
    Dim betweenEquations As Boolean
    Dim previousEnd As Long
    Dim nextStart As Long
    Dim positionInEquation As Long
    Dim equationLength As Long
    
    ' Count the number of OMath objects in the document
    oMathCount = ActiveDocument.OMaths.Count
    
    If oMathCount = 0 Then
        MsgBox "There are no equations in the document."
        Exit Sub
    End If
    
    isOnOMath = False
    currentOMathIndex = -1
    beforeFirst = False
    afterLast = False
    betweenEquations = False
    
    ' Get the selection start and end positions
    selStart = Selection.Range.Start
    selEnd = Selection.Range.End
    
    ' Check if before first equation
    If selEnd <= ActiveDocument.OMaths(1).Range.Start And CheckIfSelectionIsInOMath() = False Then
        beforeFirst = True
    End If
    
    ' Check if after last equation
    If selStart >= ActiveDocument.OMaths(oMathCount).Range.End And CheckIfSelectionIsInOMath() = False Then
        afterLast = True
    End If
    
    ' Check if the cursor is within any OMath object
    For i = 1 To oMathCount
        With ActiveDocument.OMaths(i).Range
            If selStart >= .Start And selEnd <= .End And CheckIfSelectionIsInOMath() = True Then
                isOnOMath = True
                currentOMathIndex = i
                positionInEquation = selStart - .Start
                equationLength = .End - .Start
                Exit For
            End If
        End With
        
        ' Check if between equations
        If i < oMathCount Then
            previousEnd = ActiveDocument.OMaths(i).Range.End
            nextStart = ActiveDocument.OMaths(i + 1).Range.Start
            If (selStart >= previousEnd And selEnd <= nextStart) And CheckIfSelectionIsInOMath() = False Then
                betweenEquations = True
                currentOMathIndex = i
                Exit For
            End If
        End If
    Next i
    
    isOnOMath = CheckIfSelectionIsInOMath()
    
    ' Construct the appropriate message
    If isOnOMath And positionInEquation = 0 Then
        message = "You are at the start of equation #" & currentOMathIndex & " out of " & oMathCount & " equations."
    ElseIf isOnOMath And positionInEquation = equationLength Then
        message = "You are at the end of equation #" & currentOMathIndex & " out of " & oMathCount & " equations."
    ElseIf isOnOMath Then
        message = "You are at postion " & positionInEquation & " of " & equationLength & " in equation #" & currentOMathIndex & " out of " & oMathCount & " equations."
    ElseIf beforeFirst Then
        message = "You are before the first equation. There are " & oMathCount & " equations in the document."
    ElseIf afterLast Then
        message = "You are after the last equation. There are " & oMathCount & " equations in the document."
    ElseIf betweenEquations Then
        message = "You are between equations #" & currentOMathIndex & " and #" & (currentOMathIndex + 1) & _
                 " out of " & oMathCount & " equations."
    Else
        message = "There are " & oMathCount & " equations in the document."
    End If
    
    MsgBox message
End Sub

Private Sub playTickSound()
    On Error GoTo TickErrorHandler
    If blnWavFilesWork And FileExists("C:\mathtick.wav") Then
        PlaySound "C:\mathtick.wav"
    Else
        Beep
    End If
    Exit Sub
    
TickErrorHandler:
    blnWavFilesWork = False
    Beep
    Err.Clear
End Sub

Private Sub playTockSound()
    On Error GoTo TockErrorHandler
    If blnWavFilesWork And FileExists("C:\mathtock.wav") Then
        PlaySound "C:\mathtock.wav"
    Else
        Beep
    End If
    Exit Sub
    
TockErrorHandler:
    blnWavFilesWork = False
    Beep
    Err.Clear
End Sub

Private Sub PlaySound(soundFile As String)
    ' Play the sound located in the system
    Call sndPlaySound(soundFile, &H1)
End Sub

Private Function FileExists(ByVal FileName As String) As Boolean
    FileExists = Dir(FileName) <> ""
End Function

Public Sub EditMathInLinearFormatSimple()
    Dim oMath As oMath
    Dim strLinearMath As String
    
    ' Check if selection is in a math zone
    If Not (Selection.OMaths.Count = 1 And _
           InStr(1, Selection.Font.Name, "Math", vbTextCompare) > 0) Then
        MsgBox "Please place cursor within a math expression.", vbExclamation
        Exit Sub
    End If
    
    ' Get the OMath object
    Set oMath = Selection.OMaths(1)
    oMath.Linearize
    oMath.ConvertToNormalText
    
    ' Get linear format
    strLinearMath = oMath.Range.Text
    
    ' Show input box with linear math
    strLinearMath = InputBox("Edit math expression in linear format:" & vbCrLf & _
                            "(Example: a^2 + b^2 = c^2)", _
                            "Edit Math Expression", strLinearMath)
    
    ' Check if user cancelled
    If strLinearMath = "" Then Exit Sub
    
    ' Confirm changes
    If MsgBox("Update math expression to:" & vbCrLf & vbCrLf & _
              strLinearMath, vbOKCancel + vbQuestion) = vbOK Then
        ' Delete existing math
        oMath.Range.Delete
        
        ' Insert and convert new math
        Selection.TypeText Text:=strLinearMath
        Selection.OMaths.Add Range:=Selection.Range
    End If
End Sub

Public Sub EditMathWithForm()
    Dim oMath As oMath
    Dim objRange As Range
    Dim objEq As oMath
    Dim blnNewOMath As Boolean
    Dim strMathText As String
    
    ' Check if selection is in a math zone
    blnNewOMath = Not CheckIfSelectionIsInOMath()
    
    If blnNewOMath Then
        strMathText = ""
    Else
        ' Get the OMath object
        Set oMath = Selection.OMaths(1)
        oMath.Linearize
        strMathText = oMath.Range.Text
    End If
    
    ' Show form with current math
    With New frmMathEdit
        .Initialize ConvertUnicodeToMathKeywordsOptimized(strMathText)
        .Show
        
        ' If OK was clicked and we have text
        If .FormResult = vbOK And .mathText <> "" Then
            If Not blnNewOMath Then
                ' Delete existing math
                oMath.Range.Delete
            End If
            
            ' Insert and convert new math
            Set objRange = Selection.Range
            objRange.Text = .mathText
            Set objRange = Selection.OMaths.Add(objRange)
            Set objEq = objRange.OMaths(1)
            
            ' Move to end then back to start so Equation Editor appears
            Selection.SetRange Start:=objEq.Range.End, End:=objEq.Range.End
            Selection.SetRange Start:=objEq.Range.Start, End:=objEq.Range.Start
            
            ' Try to format the new math
            On Error Resume Next
                objEq.BuildUp
            On Error GoTo 0
            
        End If
    End With
End Sub

Function ConvertUnicodeToMathKeywordsOptimized(strInput As String) As String
    Static cachedMappings() As AutoCorrectMap
    Static isCached As Boolean
    
    Dim oAutoCorrect As OMathAutoCorrect
    Dim oEntry As OMathAutoCorrectEntry
    Dim strResult As String
    Dim i As Long, j As Long
    Dim strChar As String
    
    ' Build cache if needed
    If Not isCached Then
        Set oAutoCorrect = Application.OMathAutoCorrect
        ReDim cachedMappings(1 To oAutoCorrect.Entries.Count)
        
        For i = 1 To oAutoCorrect.Entries.Count
            Set oEntry = oAutoCorrect.Entries.Item(i)
            cachedMappings(i).UnicodeChar = oEntry.Value
            cachedMappings(i).Keyword = oEntry.Name & " "
        Next i
        
        isCached = True
    End If
    
    ' Start with input string
    strResult = strInput
    
    ' Process each mapping
    For i = 1 To UBound(cachedMappings)
        With cachedMappings(i)
            If InStr(strResult, .UnicodeChar) > 0 Then
                strResult = Replace(strResult, .UnicodeChar, .Keyword)
            End If
        End With
    Next i
    
    ConvertUnicodeToMathKeywordsOptimized = strResult
End Function
