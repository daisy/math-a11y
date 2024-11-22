VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMathEdit 
   Caption         =   "Edit Equation"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "frmMathEdit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMathEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strResult As String
Private intDialogResult As VbMsgBoxResult

' Properties to get results
Public Property Get mathText() As String
    mathText = strResult
End Property

Public Property Get FormResult() As VbMsgBoxResult
    FormResult = intDialogResult
End Property

' Initialize with current math text
Public Sub Initialize(strMath As String)
    strResult = strMath
End Sub

Private Sub UserForm_Initialize()
    ' Set up form defaults
    Me.Caption = "Edit Math Expression"
    
    ' Configure textbox
    With txtMath
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Font.Name = "Consolas"  ' Good monospace font for math
        .Font.Size = 11
        .Width = 400
        .Height = 200
        

    End With
    
    ' Configure buttons
    With cmdOK
        .Caption = "OK"
        .Default = True
        .Top = txtMath.Top + txtMath.Height + 10
        .Left = txtMath.Width - .Width - cmdCancel.Width - 5
    End With
    
    With cmdCancel
        .Caption = "Cancel"
        .Cancel = True
        .Top = cmdOK.Top
        .Left = txtMath.Width - .Width
    End With
    
    ' Size form to fit controls with padding
    Me.Width = txtMath.Width + 30
    Me.Height = cmdOK.Top + cmdOK.Height + 40
    
    ' Center the form
    Me.StartUpPosition = 0 ' Manual
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub

Private Sub UserForm_Activate()
    ' Add the current math expression at cursor position
    txtMath.selStart = Len(txtMath.Text)
    txtMath.SelText = strResult
    
    ' Select the added text
    txtMath.selStart = Len(txtMath.Text) - Len(strResult)
    txtMath.SelLength = Len(strResult)
End Sub

Private Sub cmdOK_Click()

        strResult = txtMath.Text
   
    
    intDialogResult = vbOK
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    intDialogResult = vbCancel
    Me.Hide
End Sub

