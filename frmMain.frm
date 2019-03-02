VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Date Picker Sample"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetDate 
      Caption         =   "Get Date"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       frmMain
' FILENAME:     frmMain.frm
' AUTHOR:       Chris Gallucci
'*******************************************************************************
Option Explicit

'*******************************************************************************
' cmdGetDate_Click (SUB)
'
' PARAMETERS:
' None
'
' DESCRIPTION:
' Example of how to call the form-based date picker control.
'*******************************************************************************
Private Sub cmdGetDate_Click()
    Dim myDate As Date
    
    ' get date value to initialiaze the control with
    If IsDate(txtDate.Text) Then
        myDate = CDate(txtDate.Text)
    Else
        myDate = Now()
    End If
    
    ' call date picker control
    If dlgDatePicker.GetDate(myDate) Then
        ' if user didn't cancel, then set the value to the text box
        txtDate.Text = Format$(myDate, "mm/dd/yyyy")
    End If
End Sub
