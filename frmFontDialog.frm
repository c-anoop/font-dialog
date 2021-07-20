VERSION 5.00
Begin VB.Form frmFontDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Font Dialog"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   Icon            =   "frmFontDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cmbScript 
      Height          =   315
      ItemData        =   "frmFontDialog.frx":030A
      Left            =   2520
      List            =   "frmFontDialog.frx":0311
      TabIndex        =   9
      Text            =   "Western"
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Frame fraSample 
      Caption         =   "Sample"
      Height          =   1095
      Left            =   2400
      TabIndex        =   16
      Top             =   3000
      Width           =   2415
      Begin VB.TextBox txtSample 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   615
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "AaBbYyZz"
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fraEffects 
      Caption         =   "Effects"
      Height          =   1815
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   2175
      Begin VB.CheckBox chkUnderline 
         Caption         =   "Underline"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkStrikeout 
         Caption         =   "Strikeout"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox cmbSize 
      Height          =   315
      ItemData        =   "frmFontDialog.frx":031E
      Left            =   3720
      List            =   "frmFontDialog.frx":0349
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame fraStyle 
      Caption         =   "Font Style:"
      Height          =   2415
      Left            =   1920
      TabIndex        =   11
      Top             =   240
      Width           =   1695
      Begin VB.OptionButton optBoldItalic 
         Caption         =   "Bold Italic"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton optItalic 
         Caption         =   "Italic"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton optBold 
         Caption         =   "Bold"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optRegular 
         Caption         =   "Regular"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.TextBox txtFont 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.ListBox lstFont 
      Height          =   1815
      ItemData        =   "frmFontDialog.frx":0380
      Left            =   120
      List            =   "frmFontDialog.frx":03BD
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Script:"
      Height          =   195
      Left            =   2520
      TabIndex        =   18
      Top             =   4200
      Width           =   450
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Size:"
      Height          =   195
      Left            =   3720
      TabIndex        =   14
      Top             =   240
      Width           =   345
   End
   Begin VB.Label lblFont 
      AutoSize        =   -1  'True
      Caption         =   "Font:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "frmFontDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkStrikeout_Click()
  
 ' This subroutine is for strikethrough text
   txtSample.Font.Strikethrough = chkStrikeout.Value

End Sub

Private Sub chkUnderline_Click()
 
 ' This subroutine is foe underline text
   txtSample.Font.Underline = chkUnderline.Value

End Sub

Private Sub cmbSize_Click()
 
 ' This subroutine sets the size of the text
   txtSample.Font.Size = cmbSize.List _
                         (cmbSize.ListIndex)

End Sub

Private Sub cmdCancel_Click()

 ' This subroutine unloads the form
   Unload Me

End Sub

Private Sub cmdOk_Click()
 ' This subroutine unloads the form
   Unload Me

End Sub

Private Sub Form_Load()
   
 ' This subroutine sets the initial values
   optRegular.Value = True
   cmbSize.Text = "10"
   txtFont.Text = "Arial"
   txtSample.Font.Size = 10
   txtSample.Font.Name = txtFont.Text

End Sub

Private Sub lstFont_Click()
  
 ' This subroutine sets the text font as it is selected _
   in the font List box
   txtFont.Text = lstFont.List(lstFont.ListIndex)
   txtSample.Font.Name = txtFont.Text

End Sub

Private Sub optBold_Click()
 
 ' This subroutine sets the text as bold
   txtSample.Font.Bold = True
   txtSample.Font.Italic = False
   
End Sub

Private Sub optBoldItalic_Click()
   
 ' This subroutine sets the text as bold italic
   txtSample.Font.Bold = True
   txtSample.Font.Italic = True
   
End Sub

Private Sub optItalic_Click()
   
 ' This subroutine sets text as italic
   txtSample.Font.Bold = False
   txtSample.Font.Italic = True
   
End Sub

Private Sub optRegular_Click()
   
 ' This subroutine sets the text as regular
   txtSample.Font.Bold = False
   txtSample.Font.Italic = False
   
End Sub

