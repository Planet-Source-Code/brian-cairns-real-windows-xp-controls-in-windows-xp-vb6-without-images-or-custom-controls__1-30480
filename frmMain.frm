VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Windows XP Style Controls - Sample"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbDemo 
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Text            =   "Combo Boxes"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   3000
      TabIndex        =   6
      Text            =   "Text Boxes"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   4080
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command Buttons"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.OptionButton optDemo 
      Caption         =   "&Radio Buttons"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.OptionButton optDemo 
      Caption         =   "&Radio Buttons"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.CheckBox ChkDemo 
      Caption         =   "&Checkboxes"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Scroll Bars"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Menu m1 
      Caption         =   "Menus"
      Begin VB.Menu m2 
         Caption         =   "Menu Item"
      End
      Begin VB.Menu m3 
         Caption         =   "Menu Item"
      End
      Begin VB.Menu m4 
         Caption         =   "Submenu"
         Begin VB.Menu m5 
            Caption         =   "Submenu Item"
         End
         Begin VB.Menu m6 
            Caption         =   "Submenu Item"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
    Dim comctls As INITCOMMONCONTROLSEX_TYPE  ' identifies the control to register
    Dim retval As Long                        ' generic return value
    With comctls
        .dwSize = Len(comctls)
        .dwICC = ICC_INTERNET_CLASSES
    End With
    retval = InitCommonControlsEx(comctls)
End Sub

