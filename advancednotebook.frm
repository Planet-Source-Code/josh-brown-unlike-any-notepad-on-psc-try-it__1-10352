VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form form1 
   Caption         =   "Right with 2 colors or sizes with a click"
   ClientHeight    =   4815
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6480
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Left            =   480
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   720
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6376
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"2colors.frx":0000
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuopen 
         Caption         =   "Open File"
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save File"
      End
      Begin VB.Menu mnuspace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuspace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Picture1_Click()
Dim WhatColor As Long
  CommonDialog1.ShowColor
  CommonDialog1.Flags = cdlCCRGBInit
  WhatColor = CommonDialog1.Color
  Picture1.BackColor = CommonDialog1.Color
End Sub

Private Sub Picture2_Click()
Dim WhatColor As Long
  CommonDialog1.ShowColor
  CommonDialog1.Flags = cdlCCRGBInit
  WhatColor = CommonDialog1.Color
  Picture2.BackColor = CommonDialog1.Color
End Sub
Private Sub Picture3_Click()
  Dim WhatColor As Long
  CommonDialog1.ShowColor
  CommonDialog1.Flags = cdlCCRGBInit
  WhatColor = CommonDialog1.Color
  Picture3.BackColor = CommonDialog1.Color
End Sub

Private Sub rtb1_click()
 If rtb1.SelColor = Picture2.BackColor Then
 rtb1.SelColor = Picture1.BackColor
 Else
 rtb1.SelColor = Picture2.BackColor
 End If
End Sub




