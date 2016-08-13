VERSION 5.00
Begin VB.Form frmEditor_Spell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5055
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Spell.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   429
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Index"
      Height          =   255
      Left            =   1800
      TabIndex        =   32
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Frame FraData 
      Caption         =   "Data"
      Height          =   1215
      Left            =   120
      TabIndex        =   25
      Top             =   1320
      Width           =   4815
      Begin VB.HScrollBar scrlMPReq 
         Height          =   255
         Left            =   720
         Max             =   1000
         TabIndex        =   29
         Top             =   720
         Width           =   3255
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         Left            =   720
         Max             =   255
         Min             =   1
         TabIndex        =   26
         Top             =   360
         Value           =   1
         Width           =   3255
      End
      Begin VB.Label lblMPReq 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4200
         TabIndex        =   31
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblMP 
         Caption         =   "MP"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   735
      End
      Begin VB.Label lblLevel 
         Caption         =   "Level"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   735
      End
      Begin VB.Label lblLevelReq 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   27
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame FraPic 
      Caption         =   "Spell Animation"
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   4815
      Begin VB.HScrollBar scrlFrame 
         Height          =   255
         Left            =   960
         Max             =   255
         TabIndex        =   22
         Top             =   720
         Value           =   1
         Width           =   2655
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   960
         Max             =   255
         TabIndex        =   19
         Top             =   360
         Value           =   1
         Width           =   2655
      End
      Begin VB.PictureBox picPic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4200
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblFrame 
         Caption         =   "Frame"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblFrameNum 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblSpell 
         Caption         =   "Spell"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblSpellNum 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbClassReq 
      Height          =   360
      ItemData        =   "frmEditor_Spell.frx":3332
      Left            =   120
      List            =   "frmEditor_Spell.frx":3334
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   720
      Width           =   4815
   End
   Begin VB.Frame fraGiveItem 
      Caption         =   "Give Item"
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlItemValue 
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   840
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemNum 
         Height          =   255
         Left            =   1320
         Max             =   255
         Min             =   1
         TabIndex        =   10
         Top             =   360
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label lblItemValue 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4200
         TabIndex        =   15
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Value"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblItemNum 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Item"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   6000
      Width           =   1455
   End
   Begin VB.ComboBox cmbType 
      Height          =   360
      ItemData        =   "frmEditor_Spell.frx":3336
      Left            =   120
      List            =   "frmEditor_Spell.frx":334F
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3960
      Width           =   4815
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Vitals Data"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   3
         Top             =   360
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Vital Mod"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblVitalMod 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.TextBox txtName 
      Height          =   390
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmEditor_Spell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    scrlLevelReq.Max = MAX_LEVELS
End Sub

Private Sub cmdSave_Click()

    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        Call SpellEditorOk
    End If

End Sub

Private Sub cmdBack_Click()
    Unload Me
    frmIndex.Show
End Sub

Private Sub cmdCancel_Click()
    Call SpellEditorCancel
End Sub

Private Sub cmbType_Click()

    If cmbType.ListIndex <> SPELL_TYPE_GIVEITEM Then
        fraVitals.Visible = True
        fraGiveItem.Visible = False
    Else
        fraVitals.Visible = False
        fraGiveItem.Visible = True
    End If

End Sub

Private Sub scrlFrame_Change()
    lblFrameNum.Caption = scrlFrame.Value
End Sub

Private Sub scrlItemNum_Change()
    fraGiveItem.Caption = "Give Item " & Trim$(Item(scrlItemNum.Value).Name)
    lblItemNum.Caption = CStr(scrlItemNum.Value)
End Sub

Private Sub scrlItemValue_Change()
    lblItemValue.Caption = CStr(scrlItemValue.Value)
End Sub

Private Sub scrlLevelReq_Change()
    lblLevelReq.Caption = CStr(scrlLevelReq.Value)
End Sub

Private Sub scrlMPReq_Change()
    lblMPReq.Caption = CStr(scrlMPReq.Value)
End Sub

Private Sub scrlPic_Change()
    lblSpellNum = CStr(scrlPic.Value)
    frmEditor_Spell.scrlFrame.Value = 0
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = CStr(scrlVitalMod.Value)
End Sub
