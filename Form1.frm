VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Beep"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3210
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   3210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWarble 
      Caption         =   "Warble"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdStartSound 
      Caption         =   "Beep"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtLength 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Text            =   "1000"
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtFreq 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "1000"
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Internal Speaker Beep Routine
'by Bob Richards
'adapted from Jorge Loubet's code for Win9x
'added code for NT2000
'
'This program tests for the operating system and executes
'a beep routine appropriate for the system.
'
'For Win9x systems, the file WIN95IO.DLL must be copied
'to the Windows/System folder.
'WIN95IO.DLL is available from http://www.softcircuits.com

Private Sub cmdStartSound_Click()
    Dim FreqHz As Single
    Dim LengthMs As Single
    
    FreqHz = Val(txtFreq.Text)     'In Hz
    LengthMs = Val(txtLength.Text) 'In ms
    
    PcSpeakerBeep FreqHz, LengthMs

End Sub


Private Sub cmdWarble_Click()
    Dim LengthMs As Single
    Dim FreqHz As Integer
    
    LengthMs = Val(txtLength.Text) 'In ms
    FreqHz = Val(txtFreq.Text)
    Call Warble(FreqHz, LengthMs)
End Sub

Private Sub Form_Load()

    'Just to show the user which operating system is being used
    MsgBox GetVersion
           
    
End Sub

