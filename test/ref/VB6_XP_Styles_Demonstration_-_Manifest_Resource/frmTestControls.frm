VERSION 5.00
Begin VB.Form frmTestControls 
   Caption         =   "XP Visual Styles Tester - Manifest Stored in Resource"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8895
   Icon            =   "frmTestControls.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTest 
      AutoSize        =   -1  'True
      Height          =   1260
      Left            =   7260
      Picture         =   "frmTestControls.frx":1272
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   24
      Top             =   4380
      Width           =   1260
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Advanced..."
      Height          =   435
      Left            =   4380
      TabIndex        =   22
      Top             =   6660
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7440
      TabIndex        =   20
      Top             =   6660
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5940
      TabIndex        =   19
      Top             =   6660
      Width           =   1395
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1815
      Left            =   8580
      TabIndex        =   18
      Top             =   4380
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   6240
      Width           =   6975
   End
   Begin VB.ComboBox cboDropDownList 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmTestControls.frx":1ACA
      Left            =   1560
      List            =   "frmTestControls.frx":1ADD
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2340
      Width           =   7275
   End
   Begin VB.ListBox lstTest 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      ItemData        =   "frmTestControls.frx":1B54
      Left            =   1560
      List            =   "frmTestControls.frx":1B73
      TabIndex        =   12
      Top             =   2700
      Width           =   7275
   End
   Begin VB.ComboBox cboDropDownCombo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmTestControls.frx":1C19
      Left            =   1560
      List            =   "frmTestControls.frx":1C26
      TabIndex        =   11
      Text            =   "Drop-Down Combo"
      Top             =   1980
      Width           =   7275
   End
   Begin VB.Frame fraTestOption 
      Caption         =   "Frame - Option Boxes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3720
      TabIndex        =   10
      Top             =   4440
      Width           =   2115
      Begin VB.PictureBox picOptions 
         BorderStyle     =   0  'None
         Height          =   1395
         Left            =   60
         ScaleHeight     =   1395
         ScaleWidth      =   1995
         TabIndex        =   25
         Top             =   240
         Width           =   1995
         Begin VB.OptionButton optTest 
            Caption         =   "Option Item 1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   0
            Left            =   60
            MaskColor       =   &H80000008&
            TabIndex        =   29
            Top             =   60
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optTest 
            Caption         =   "Option Item 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   1
            Left            =   60
            MaskColor       =   &H8000000F&
            TabIndex        =   28
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optTest 
            Caption         =   "Option Item 3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   2
            Left            =   60
            MaskColor       =   &H8000000F&
            TabIndex        =   27
            Top             =   660
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optTest 
            Caption         =   "Option Item 4"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   3
            Left            =   60
            MaskColor       =   &H8000000F&
            TabIndex        =   26
            Top             =   960
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
      End
   End
   Begin VB.Frame fraTest 
      Caption         =   "Frame - Check boxes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1560
      TabIndex        =   5
      Top             =   4440
      Width           =   2115
      Begin VB.CheckBox chkTest 
         Caption         =   "Check Option 4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkTest 
         Caption         =   "Check Option 3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   900
         Width           =   1695
      End
      Begin VB.CheckBox chkTest 
         Caption         =   "Check Option 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkTest 
         Caption         =   "Check Option 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   1695
      End
   End
   Begin VB.TextBox txtMultiLine 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "frmTestControls.frx":1C60
      Top             =   960
      Width           =   7275
   End
   Begin VB.TextBox txtSingleLine 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Text            =   "Test of a single line text box"
      Top             =   600
      Width           =   7275
   End
   Begin VB.Image imgTest 
      BorderStyle     =   1  'Fixed Single
      Height          =   1275
      Left            =   5940
      Top             =   4380
      Width           =   1275
   End
   Begin VB.Label lblInfo 
      Caption         =   "Buttons:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   23
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      Caption         =   "Frames, Scroll bars, Check boxes and Option Boxes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   5
      Left            =   60
      TabIndex        =   21
      Top             =   4500
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Caption         =   "List Box:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   16
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      Caption         =   "Drop-down Combo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   15
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Caption         =   "Drop-down List:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   14
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      Caption         =   "Multi-line Text:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   1020
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      Caption         =   "Single-line Text:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   1335
   End
   Begin VB.Label llbInfo 
      Caption         =   "This form simply contains a number of controls to test XP Visual styles against."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   8715
   End
End
Attribute VB_Name = "frmTestControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   imgTest.Picture = picTest.Picture
End Sub
