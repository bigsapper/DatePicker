VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form dlgDatePicker 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2940
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   2715
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "dlgDatePicker.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      Default         =   -1  'True
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   1560
      TabIndex        =   1
      Top             =   2520
      Width           =   1092
   End
   Begin MSComCtl2.MonthView cal 
      Height          =   2370
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   22740993
      CurrentDate     =   37620
   End
End
Attribute VB_Name = "dlgDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       dlgDatePicker
' FILENAME:     dlgDatePicker.frm
' AUTHOR:       Chris Gallucci
' COPYRIGHT:    Copyright 2003 Chris Gallucci. All Rights Reserved.
'               http://www.dotnetconsultant.com
'               Author grants royalty-free rights to use this code within
'               compiled applications. Selling or otherwise distributing
'               this source code is not allowed without author's express
'               permission.
'*******************************************************************************
'
' DESCRIPTION:
' Form-based Date Picker control using the standard Microsoft Calendar control.
' Created using Microsoft Visual Basic 6 (SP5).
'
'*******************************************************************************
Option Explicit

' private variables
Private m_CurrDate As Date
Private m_AcceptChange As Boolean

'*******************************************************************************
' GetDate (FUNCTION)
'
' PARAMETERS:
' (In/Out) - UserDate - Date - Value used to initialize the control and also
' contains the return value that's selected when the function returns.
'
' RETURN VALUE:
' Boolean - Indicates whether the user canceled the selection of a date value.
'
' DESCRIPTION:
'  Primary function that displays the control when called and returns the
'  selected date.
'*******************************************************************************
Public Function GetDate(UserDate As Date) As Boolean
    ' store user-specified date
    m_CurrDate = UserDate
    cal.Value = m_CurrDate
    
    ' display this form
    Me.Show vbModal
    
    ' return selected date
    If m_AcceptChange Then
        UserDate = m_CurrDate
    End If
    
    ' return value indicates if date was selected
    GetDate = m_AcceptChange
End Function

Private Sub cal_DateClick(ByVal DateClicked As Date)
    m_CurrDate = DateClicked
    cmdOK.SetFocus
End Sub

Private Sub cal_DateDblClick(ByVal DateDblClicked As Date)
    m_CurrDate = DateDblClicked
    m_AcceptChange = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    m_AcceptChange = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    m_AcceptChange = True
    Unload Me
End Sub

Private Sub Form_Load()
    ' set default date
    cal.Value = Now
    m_AcceptChange = False
End Sub
