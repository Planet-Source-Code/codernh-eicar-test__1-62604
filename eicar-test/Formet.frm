VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Formet 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Eicar Test"
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4260
   Icon            =   "Formet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Formet.frx":030A
   ScaleHeight     =   4845
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H0094B68C&
      Caption         =   "Whats the String?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      Picture         =   "Formet.frx":43E8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "X5O!P%@AP[4\PZX54(P^)7CC)7}$EICAR-STANDARD-ANTIVIRUS-TEST-FILE!$H+H*"
      Top             =   720
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0094B68C&
      Caption         =   "TEST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Picture         =   "Formet.frx":46F2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Formet.frx":49FC
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EICAR Anti-Virus TESTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0094B68C&
      BorderWidth     =   4
      X1              =   0
      X2              =   4200
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3720
      Picture         =   "Formet.frx":4ADB
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Formet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'EICAR ( European Institute for Computer Anti-Virus Research) is a highly
'effective way of testing anti virus software..the text string is 68 characters long
'and HAS to be 68 bytes in size...the EICAR test virus is NOT really a virus. It cannot do any harm to your system, even if your software does not detect it.
'It is supported by most leading anti-virus software makers... such as IBM, McAfee, Sophos, and Symantec/Norton.
'i made this as an easy way to test...DO NOT change text file name when
'dialog comes up...it has to be eicar.com you can learn more about eicar by visiting
'http://www.eicar.org/ ... i made this for anyone considering making an anti-virus software
'in visual basic and wanted an onboard tester
Private Sub Command1_Click()
MsgBox "save the text file to a folder you can access....if you have an effective anti-virus it should prompt you to delete it..but if it does not..you can manually delete it as it is harmless.REMEMBER NOT TO CHANGE FILE NAME!!!!"
    Dim TestString, X As Integer
    On Error GoTo CloseIt
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
    CommonDialog1.FileName = "eicar.com"
    CommonDialog1.FilterIndex = 0
    CommonDialog1.ShowSave
    TestString = CommonDialog1.FileName
    If Len(Dir$(TestString)) <> 0 Then
        X = MsgBox("This file already exists: " + TestString + ", do you want replace it?", vbYesNo + vbCritical, "Error")
        If X = vbNo Then Exit Sub
    End If
    Open TestString For Output As #1
        Print #1, Text1.Text
    Close #1
CloseIt:     Exit Sub

End Sub


Private Sub Command2_Click()
MsgBox "X5O!P%@AP[4\PZX54(P^)7CC)7}$EICAR-STANDARD-ANTIVIRUS-TEST-FILE!$H+H*       <----that is the exact EICAR test string."
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
Call SendMessage(Me.hwnd, &HA1, 2, 0&)

End Sub

Private Sub Image1_Click()
End
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
Call SendMessage(Me.hwnd, &HA1, 2, 0&)

End Sub


Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
Call SendMessage(Me.hwnd, &HA1, 2, 0&)

End Sub
