VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�_���m�ѧO��  ��"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4350
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   162
   ScaleMode       =   3  '����
   ScaleWidth      =   290
   Begin �_���m�ѧO��.PlowPanel ColorInfo 
      Height          =   2415
      Left            =   2535
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   15
      Width           =   1815
      _extentx        =   3201
      _extenty        =   4260
      Begin �_���m�ѧO��.PlowCButton SColor 
         Height          =   285
         Left            =   1155
         TabIndex        =   64
         TabStop         =   0   'False
         ToolTipText     =   "���@�U�H��ܦ��w����m"
         Top             =   445
         Width           =   540
         _extentx        =   953
         _extenty        =   503
         backcolor       =   0
         backcolor       =   0
      End
      Begin �_���m�ѧO��.PlowCButton NColor 
         Height          =   285
         Left            =   1155
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "���@�U�H��ܦ������m"
         Top             =   120
         Width           =   540
         _extentx        =   953
         _extenty        =   503
         backcolor       =   0
         backcolor       =   0
      End
      Begin �_���m�ѧO��.PlowCButton TColor 
         Height          =   615
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "���@�U�H�˵���h��T"
         Top             =   120
         Width           =   990
         _extentx        =   1746
         _extenty        =   1085
         backcolor       =   0
         backcolor       =   0
      End
      Begin �_���m�ѧO��.PlowFlatButton Info 
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   840
         Width           =   1575
         _extentx        =   2778
         _extenty        =   423
         caption         =   "�зǦW�Gblack"
         font            =   "Main.frx":0442
      End
      Begin �_���m�ѧO��.PlowFlatButton Info 
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "���@�U�H�ƻs��ŶKï�]��i�b�����m���U Ctrl-C �H�ƻs�^"
         Top             =   1320
         Width           =   1575
         _extentx        =   2778
         _extenty        =   423
         caption         =   "�зǭȡG#000000"
         font            =   "Main.frx":046A
      End
      Begin �_���m�ѧO��.PlowFlatButton Info 
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1575
         _extentx        =   2778
         _extenty        =   423
         caption         =   "²�g�ȡG#000"
         font            =   "Main.frx":0492
      End
      Begin �_���m�ѧO��.PlowFlatButton Info 
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "���@�U�H�ƻs��ŶKï"
         Top             =   1800
         Width           =   1575
         _extentx        =   2778
         _extenty        =   423
         caption         =   "����ȡG#000"
         font            =   "Main.frx":04BA
      End
      Begin �_���m�ѧO��.PlowFlatButton Info 
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "���@�U�H�ƻs��ŶKï"
         Top             =   2040
         Width           =   1575
         _extentx        =   2778
         _extenty        =   423
         caption         =   "�w���ȡG#000"
         font            =   "Main.frx":04E2
      End
      Begin �_���m�ѧO��.PlowFlatButton Info 
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1575
         _extentx        =   2778
         _extenty        =   423
         caption         =   "�����W�Gblack"
         font            =   "Main.frx":050A
      End
      Begin VB.TextBox hiddenText1 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.TextBox hiddenText2 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2760
      TabIndex        =   67
      Text            =   "Text1"
      Top             =   1200
      Width           =   615
   End
   Begin �_���m�ѧO��.PlowPanel Fra 
      Height          =   2415
      Index           =   3
      Left            =   0
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      Width           =   2535
      _extentx        =   4471
      _extenty        =   4260
      Begin �_���m�ѧO��.PlowFlatButton ScrStart 
         Height          =   255
         Left            =   120
         TabIndex        =   72
         ToolTipText     =   "���@�U�H�}�l�^���ù���m"
         Top             =   120
         Width           =   435
         _extentx        =   767
         _extenty        =   450
         caption         =   "�}�l"
         font            =   "Main.frx":0532
      End
      Begin VB.PictureBox picScr 
         Appearance      =   0  '����
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   120
         ScaleHeight     =   119
         ScaleMode       =   3  '����
         ScaleWidth      =   151
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "���@�U�H��ܦ�m"
         Top             =   480
         Width           =   2295
         Begin VB.Timer TimerScr 
            Enabled         =   0   'False
            Interval        =   200
            Left            =   120
            Top             =   240
         End
      End
      Begin VB.Label lblPos 
         Caption         =   "�y�СG"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   59
         Top             =   150
         Width           =   1695
      End
   End
   Begin �_���m�ѧO��.PlowPanel Fra 
      Height          =   2415
      Index           =   6
      Left            =   0
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      Width           =   2535
      _extentx        =   4471
      _extenty        =   4260
      Begin VB.PictureBox SizP 
         Appearance      =   0  '����
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   68
         Top             =   2040
         Width           =   255
         Begin VB.Image Siz 
            Height          =   180
            Left            =   30
            Picture         =   "Main.frx":055A
            ToolTipText     =   "���@�U�H��ܭ��v�A�ΥH�ƹ��u���ֳt����"
            Top             =   30
            Width           =   180
         End
      End
      Begin VB.HScrollBar Hs 
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2040
         Width           =   2040
      End
      Begin VB.VScrollBar Vs 
         Enabled         =   0   'False
         Height          =   1920
         Left            =   2160
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox PP 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         ScaleHeight     =   125
         ScaleMode       =   3  '����
         ScaleWidth      =   133
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "�H�����ܦ�m�A�H�k��즲�Ϥ���m"
         Top             =   120
         Width           =   2055
         Begin VB.Timer Timer 
            Interval        =   20
            Left            =   120
            Top             =   120
         End
      End
   End
   Begin �_���m�ѧO��.PlowPanel Fra 
      Height          =   2415
      Index           =   4
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      Width           =   2535
      _extentx        =   4471
      _extenty        =   4260
      Begin VB.PictureBox PK 
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2160
         Left            =   90
         ScaleHeight     =   2160
         ScaleWidth      =   2055
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   120
         Width           =   2055
         Begin VB.PictureBox KP 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  '�S���ؽu
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   12900
            Left            =   0
            ScaleHeight     =   12900
            ScaleWidth      =   2055
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   2055
            Begin VB.PictureBox SK 
               Appearance      =   0  '����
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000000&
               BorderStyle     =   0  '�S���ؽu
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Index           =   2
               Left            =   0
               ScaleHeight     =   495
               ScaleWidth      =   2055
               TabIndex        =   25
               TabStop         =   0   'False
               ToolTipText     =   "���@�U�H��ܦ�m"
               Top             =   1920
               Width           =   2055
            End
            Begin VB.PictureBox SK 
               Appearance      =   0  '����
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000000&
               BorderStyle     =   0  '�S���ؽu
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Index           =   1
               Left            =   0
               ScaleHeight     =   495
               ScaleWidth      =   2055
               TabIndex        =   24
               TabStop         =   0   'False
               ToolTipText     =   "���@�U�H��ܦ�m"
               Top             =   1080
               Width           =   2055
            End
            Begin VB.PictureBox SK 
               Appearance      =   0  '����
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000000&
               BorderStyle     =   0  '�S���ؽu
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Index           =   0
               Left            =   0
               ScaleHeight     =   495
               ScaleWidth      =   2055
               TabIndex        =   20
               TabStop         =   0   'False
               ToolTipText     =   "���@�U�H��ܦ�m"
               Top             =   240
               Width           =   2055
            End
            Begin VB.Label KL 
               AutoSize        =   -1  'True
               Caption         =   "HTML �зǦW�٦�L�G"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   23
               Top             =   0
               Width           =   1785
            End
            Begin VB.Label KL 
               AutoSize        =   -1  'True
               Caption         =   "�����W�٦�L�G"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   1
               Left            =   0
               TabIndex        =   22
               Top             =   840
               Width           =   1260
            End
            Begin VB.Label KL 
               AutoSize        =   -1  'True
               Caption         =   "�󥭥x�w����L�G"
               BeginProperty Font 
                  Name            =   "�s�ө���"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   2
               Left            =   0
               TabIndex        =   21
               Top             =   1680
               Width           =   1440
            End
         End
      End
      Begin VB.VScrollBar KVS 
         Height          =   2160
         Left            =   2190
         SmallChange     =   10
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   120
         Width           =   255
      End
   End
   Begin �_���m�ѧO��.PlowPanel Fra 
      Height          =   2415
      Index           =   2
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      Width           =   2535
      _extentx        =   4471
      _extenty        =   4260
      Begin �_���m�ѧO��.PlowColorH IColor 
         Height          =   1815
         Left            =   2040
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "���@�U�ΥH�ƹ��u���H�M�w���ס]�ΫG�ס^"
         Top             =   120
         Width           =   375
         _extentx        =   661
         _extenty        =   3201
      End
      Begin VB.OptionButton isHSV 
         Caption         =   "HSV �t��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "��� HSV ��L"
         Top             =   2040
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton isHSL 
         Caption         =   "HSL �t��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "��� HSL ��L"
         Top             =   2040
         Width           =   1095
      End
      Begin �_���m�ѧO��.PlowColorB HSColor 
         Height          =   1815
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "���@�U�ΥH��V��H�M�w��۩M�m��"
         Top             =   120
         Width           =   1815
         _extentx        =   3201
         _extenty        =   3201
         style           =   "HS"
      End
   End
   Begin �_���m�ѧO��.PlowPanel Fra 
      Height          =   2415
      Index           =   0
      Left            =   0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      Width           =   2535
      _extentx        =   4471
      _extenty        =   4260
      Begin �_���m�ѧO��.PlowHand Hand 
         Height          =   2175
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   2295
         _extentx        =   4048
         _extenty        =   3836
      End
   End
   Begin �_���m�ѧO��.PlowPanel Fra 
      Height          =   2415
      Index           =   5
      Left            =   0
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      Width           =   2535
      _extentx        =   4471
      _extenty        =   4260
      Begin VB.PictureBox PS 
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2160
         Left            =   120
         ScaleHeight     =   2160
         ScaleWidth      =   2055
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   135
         Width           =   2055
         Begin VB.PictureBox SP 
            BorderStyle     =   0  '�S���ؽu
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6720
            Left            =   0
            ScaleHeight     =   6720
            ScaleWidth      =   2055
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   0
            Width           =   2055
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   0
               Left            =   0
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   0
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "�{�ε������ؽu"
               font            =   "Main.frx":05CD
               backcolor       =   -2147483638
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   4
               Left            =   0
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   960
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "���s��"
               font            =   "Main.frx":05F5
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   23
               Left            =   0
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   5520
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "���骺���L���v"
               font            =   "Main.frx":061D
               backcolor       =   -2147483626
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   12
               Left            =   0
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   2880
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "�D�{�ε������ؽu"
               font            =   "Main.frx":0645
               backcolor       =   -2147483637
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   16
               Left            =   0
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   3840
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               forecolor       =   -2147483624
               caption         =   "�u�㴣�ܤ�r"
               font            =   "Main.frx":066D
               backcolor       =   -2147483625
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   20
               Left            =   0
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   4800
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               forecolor       =   -2147483633
               caption         =   "���骺���`���v"
               font            =   "Main.frx":0695
               backcolor       =   -2147483627
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   24
               Left            =   0
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   5760
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "���鳱�v"
               font            =   "Main.frx":06BD
               backcolor       =   -2147483632
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   1
               Left            =   0
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   240
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               forecolor       =   -2147483639
               caption         =   "�{�ε������D�C"
               font            =   "Main.frx":06E5
               backcolor       =   -2147483646
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   5
               Left            =   0
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   1200
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "���s���ϥծĪG"
               font            =   "Main.frx":070D
               backcolor       =   -2147483628
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   9
               Left            =   0
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   2160
               Width           =   2040
               _extentx        =   3387
               _extenty        =   423
               caption         =   "�Ȥ�@�Ϊ���r"
               font            =   "Main.frx":0735
               backcolor       =   -2147483631
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   13
               Left            =   0
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   3120
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               forecolor       =   -2147483629
               caption         =   "�D�{�ε������D�C"
               font            =   "Main.frx":075D
               backcolor       =   -2147483645
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   17
               Left            =   0
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   4080
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "�\���C"
               font            =   "Main.frx":0785
               backcolor       =   -2147483644
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   21
               Left            =   0
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   5040
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "�����"
               font            =   "Main.frx":07AD
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   25
               Left            =   0
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   6000
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "�����I����m"
               font            =   "Main.frx":07D5
               backcolor       =   -2147483643
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   2
               Left            =   0
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   480
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "���ε{���u�@��"
               font            =   "Main.frx":07FD
               backcolor       =   -2147483636
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   6
               Left            =   0
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   1440
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "���s�����v"
               font            =   "Main.frx":0825
               backcolor       =   -2147483632
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   10
               Left            =   0
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   2400
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               forecolor       =   -2147483634
               caption         =   "�ϥ����"
               font            =   "Main.frx":084D
               backcolor       =   -2147483635
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   14
               Left            =   0
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   3360
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "�D�{�ε������D�C��r"
               font            =   "Main.frx":0875
               backcolor       =   -2147483629
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   18
               Left            =   0
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   4320
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               forecolor       =   -2147483644
               caption         =   "�\����r"
               font            =   "Main.frx":089D
               backcolor       =   -2147483641
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   22
               Left            =   0
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   5280
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "����ϥծĪG"
               font            =   "Main.frx":08C5
               backcolor       =   -2147483628
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   26
               Left            =   0
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   6240
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               forecolor       =   -2147483633
               caption         =   "�����ج["
               font            =   "Main.frx":08ED
               backcolor       =   -2147483642
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   3
               Left            =   0
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   720
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "�ୱ"
               font            =   "Main.frx":0915
               backcolor       =   -2147483647
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   7
               Left            =   0
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   1680
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               forecolor       =   -2147483628
               caption         =   "���s�W����r"
               font            =   "Main.frx":093D
               backcolor       =   -2147483630
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   11
               Left            =   0
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   2640
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               forecolor       =   -2147483635
               caption         =   "�ϥ���ܪ���r"
               font            =   "Main.frx":0965
               backcolor       =   -2147483634
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   15
               Left            =   0
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   3600
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               forecolor       =   -2147483625
               caption         =   "�u�㴣��"
               font            =   "Main.frx":098D
               backcolor       =   -2147483624
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   19
               Left            =   0
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   4560
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "���b"
               font            =   "Main.frx":09B5
               backcolor       =   -2147483648
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   8
               Left            =   0
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   1920
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               caption         =   "�{�ε������D�C����r"
               font            =   "Main.frx":09DD
               backcolor       =   -2147483639
            End
            Begin �_���m�ѧO��.PlowFlatButton Sy 
               Height          =   240
               Index           =   27
               Left            =   0
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   6480
               Width           =   2040
               _extentx        =   3598
               _extenty        =   423
               forecolor       =   -2147483643
               caption         =   "������r"
               font            =   "Main.frx":0A05
               backcolor       =   -2147483640
            End
         End
      End
      Begin VB.VScrollBar SVS 
         Height          =   2160
         LargeChange     =   128
         Left            =   2190
         Max             =   288
         SmallChange     =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   120
         Width           =   255
      End
   End
   Begin �_���m�ѧO��.PlowPanel Fra 
      Height          =   2415
      Index           =   1
      Left            =   0
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      Width           =   2535
      _extentx        =   4471
      _extenty        =   4260
      Begin �_���m�ѧO��.PlowCHand CHand 
         Height          =   2175
         Left            =   120
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   120
         Width           =   2295
         _extentx        =   4048
         _extenty        =   3836
      End
   End
   Begin VB.Label LB 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2640
      TabIndex        =   71
      Top             =   120
      Width           =   165
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   312
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu M_Fun 
      Caption         =   "�\����(&F)"
      Begin VB.Menu MF_Fra 
         Caption         =   "��ʰt��(&R)"
         Index           =   0
      End
      Begin VB.Menu MF_Fra 
         Caption         =   "�L��t��(&P)"
         Index           =   1
      End
      Begin VB.Menu MF_Fra 
         Caption         =   "��t��L(&I)"
         Index           =   2
      End
      Begin VB.Menu MF_Fra 
         Caption         =   "�ù��ߦ�(&C)"
         Index           =   3
      End
      Begin VB.Menu MF_Fra 
         Caption         =   "������L(&S)"
         Index           =   4
      End
      Begin VB.Menu MF_Fra 
         Caption         =   "�t�ΦW��(&Y)"
         Index           =   5
      End
      Begin VB.Menu MF_L1 
         Caption         =   "-"
      End
      Begin VB.Menu MF_Pas 
         Caption         =   "�K�W��X(&V)"
      End
   End
   Begin VB.Menu M_Pic 
      Caption         =   "�Ϥ��ߦ�(&P)"
      Begin VB.Menu MP_Fil 
         Caption         =   "�q�ɮ�(&F)..."
      End
      Begin VB.Menu MP_Cli 
         Caption         =   "�q�ŶKï(&C)"
      End
   End
   Begin VB.Menu M_Opt 
      Caption         =   "�]�w(&O)"
      Begin VB.Menu MO_Top 
         Caption         =   "�̤W�h���(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu MO_Sha 
         Caption         =   "��X�]�t�u#�v(&S)"
         Checked         =   -1  'True
         Shortcut        =   ^S
      End
      Begin VB.Menu MH_L1 
         Caption         =   "-"
      End
      Begin VB.Menu MO_Abo 
         Caption         =   "����(&A)"
      End
   End
   Begin VB.Menu M_Siz 
      Caption         =   "���(&S)"
      Visible         =   0   'False
      Begin VB.Menu MS_x 
         Caption         =   "16x"
         Index           =   1
      End
      Begin VB.Menu MS_x 
         Caption         =   "8x"
         Index           =   2
      End
      Begin VB.Menu MS_x 
         Caption         =   "4x"
         Index           =   3
      End
      Begin VB.Menu MS_x 
         Caption         =   "2x"
         Index           =   4
      End
      Begin VB.Menu MS_x 
         Caption         =   "1x"
         Checked         =   -1  'True
         Index           =   5
      End
      Begin VB.Menu MS_x 
         Caption         =   "1/2x"
         Index           =   6
      End
      Begin VB.Menu MS_x 
         Caption         =   "1/4x"
         Index           =   7
      End
      Begin VB.Menu MS_x 
         Caption         =   "1/8x"
         Index           =   8
      End
      Begin VB.Menu MS_x 
         Caption         =   "1/16x"
         Index           =   9
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private origFrame As Integer                        ' �������Ϥ��ߦ⤧�e����������
Private TopMostMode As Boolean
Private LeadingSharp As Boolean

Private cImage As New c32bppDIB

Private XX As Single                                ' �ثe�����v��
Private XXIndex As Integer, XXMaxIndex As Integer   ' �O�����v�����޻P�̤j���ޭ���
Private OldX As Integer, OldY As Integer            ' �O���Ϥ��쥻��Ǫ����ߦ�m
Private PPWidth As Integer, PPHeight As Integer     ' �ثe�� PP �Y��j�p
Private PaintLock As Boolean                        ' ���ܭ��v����wø�s
Private PictureLoaded As Boolean                    ' �ХܬO�_���Ϥ��Q���J
Private PWidth As Integer, PHeight As Integer       ' �ثe�Ϥ����j�p
Private DPDC() As Long, DPBitmap() As Long          ' �Ψ��x�s�Y�Ϫ� hDC �M hBitmap �}�C
Private PPToolTipTimeout As Integer                 ' �O�� PP �����v���ܮ����ɶ�
Private PictureX As Integer, PictureY As Integer    ' �ƹ��۹���Y���Ϥ�����m
Private ControlX As Integer, ControlY As Integer    ' �ƹ��۹��Ϥ��������m
Private PictureMouseButton As Integer               ' �Y�Ϥ��ߦ�إ��Q���U�A���γ����ֳt��

Public ScrCapturing As Boolean
Private TempColor As Long

Private EventRaised As Boolean

Private Const XXCIndex = 5
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&

Private Const COLOR_BTNFACE = 15

Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetCommandLineW Lib "kernel32" () As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal Handletype As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, Optional ByVal lpOverlapped As Long = 0&) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, Optional ByVal lpOverlapped As Long = 0&) As Long







' ---------------------------------
' �{���֤�
' ---------------------------------

Private Sub Form_Load()
    Static sReadBuffer As String, lBytesRead As Long, hStdIn As Long
    Static I As Variant, TC As Long
    Static inQuote As Boolean
    
    If GetDeviceCaps(Me.hDC, 12&) < 24 Then MsgBox "���F���T���R��m�A��ĳ�z�N�ù��վ�ܥ��m�Ҧ��C"
    PrevWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf WndProc)
    EnableKBDHook
    
    ' �]�w����
    Me.Top = (Screen.Height - Me.Height) / 2: Me.Left = (Screen.Width - Me.Width) / 2
    Me.Caption = "�_���m�ѧO�� " & getAppVersion() & " ��"
    DragAcceptFiles Me.hWnd, True
    Fra(0).Visible = True: origFrame = 0
    
    ' �]�w����
    HSVMode = True: HSColor.Draw: IColor.Draw
    PictureLoaded = False
    PaintLock = False
    ScrCapturing = False
    TopMostMode = False
    LeadingSharp = True
    EventRaised = False
    PictureMouseButton = 0
    
    ' �]�w���v�M��
    XX = 1
    ReDim DPDC(MS_x.Count)
    ReDim DPBitmap(MS_x.Count)
    For Each I In MS_x
        I.Tag = 2 ^ (XXCIndex - I.Index)
        If I.Tag = XX Then XXIndex = I.Index
        If I.Tag <= 1 Then
            DPDC(I.Index) = CreateCompatibleDC(0)
        End If
    Next I
    PP.Tag = PP.ToolTipText
    PPToolTipTimeout = 0
    
    ' �]�w������L�P�t�ΦW��
    PaintWebColor
    SVS.Max = (SP.Height - PS.Height)
    SVS.LargeChange = PS.Height
    For I = 0 To 27
        Sy(I).SetToolTipText "���@�U�H�N�N�X�u" & Sys(I) & "�v�ƻs��ŶKï"
    Next I
    GetColorInfo vbBlack
    
    ' �ˬd�O�_���зǿ�J
    sReadBuffer = String(30, 0)
    hStdIn = GetStdHandle(STD_INPUT_HANDLE)
    ReadFile hStdIn, sReadBuffer, Len(sReadBuffer), lBytesRead
    sReadBuffer = Left(sReadBuffer, lBytesRead): TC = StringToColor(sReadBuffer)
    If TC <> -1 Then GetColorInfo TC
    
    ' �ˬd�O�_���R�C�C��J
    sReadBuffer = Ptr2StrU(GetCommandLineW())
    inQuote = False
    For I = 1 To Len(sReadBuffer)
        If Mid(sReadBuffer, I, 1) = """" Then inQuote = Not inQuote
        If Mid(sReadBuffer, I, 1) = " " And Not inQuote Then sReadBuffer = Right(sReadBuffer, Len(sReadBuffer) - I): Exit For
    Next I
    If Left(sReadBuffer, 1) = """" Then sReadBuffer = Mid(sReadBuffer, 2, Len(sReadBuffer) - 2)
    If sReadBuffer <> "" Then
        If TryLoadPicture(sReadBuffer) Then
            LoadcImage
        Else
            TC = StringToColor(sReadBuffer)
            If TC <> -1 Then GetColorInfo TC
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Static sWriteBuffer As String, lBytesWritten As Long, hStdOut As Long
    Static C As New ColorInfo, I As Integer
    Static lResult As Long
    
    ' �Ѱ��U�ؾ���
    lResult = SetWindowLong(Me.hWnd, GWL_WNDPROC, PrevWndProc)
    Unload FScreen
    UnHookKBD
    
    For I = 1 To MS_x.Count
        If MS_x(I).Tag <= 1 Then DeleteDC DPDC(I): DeleteObject DPBitmap(I)
    Next I
    
    ' �i��зǿ�X
    C.Color = TColor.BackColor
    sWriteBuffer = "#" & C.getRRGGBB
    hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
    WriteFile hStdOut, sWriteBuffer, Len(sWriteBuffer), lBytesWritten
End Sub








' ---------------------------------
' �������
' ---------------------------------

Private Sub M_Fun_Click()
    MF_Fra_Click origFrame
End Sub

Private Sub MF_Fra_Click(Index As Integer)
    Static K As Object
    For Each K In Fra
        If K.Index = Index Then K.Visible = True Else K.Visible = False
    Next K
    If Index = 0 Then Hand.Color = TColor.BackColor: Hand.SetTextFocus 1
    If Index = 1 Then CHand.Color = TColor.BackColor: CHand.SetTextFocus 1
    If Index = 2 Then FHS TColor.BackColor
    origFrame = Index
End Sub

Private Sub MF_Pas_Click()
    Static CC As String, SC As Long
    If Clipboard.GetFormat(1) Then
        CC = Clipboard.GetText: SC = StringToColor(CC)
        If SC <> -1 Then
            Hand.Color = SC
            CHand.Color = SC
            FHS TColor.BackColor
        Else
            If Not KBDHooked Then MsgBox "�L�k���Ѧ���X�G�u" & CC & "�v" & vbCrLf & _
                "�ШϥΦ�m�W�١B�зǭȩ�²�g�ȡC", vbCritical
        End If
    End If
End Sub

Private Sub MO_Abo_Click()
    If TopMostMode Then SetWindowPos About.hWnd, HWND_TOPMOST, _
        0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    About.Show 1
End Sub

Private Sub MO_Sha_Click()
    LeadingSharp = Not LeadingSharp
    MO_Sha.Checked = LeadingSharp
    GetColorInfo TColor.BackColor
End Sub

Private Sub MO_Top_Click()
    TopMostMode = Not TopMostMode
    MO_Top.Checked = TopMostMode
    SetWindowPos Me.hWnd, IIf(TopMostMode, HWND_TOPMOST, HWND_NOTOPMOST), _
        0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub









' ---------------------------------
' ��m��T��
' ---------------------------------

Private Sub SColor_Click()
    Hand.Color = SColor.BackColor
    CHand.Color = SColor.BackColor
    FHS SColor.BackColor
End Sub

Private Sub NColor_Click()
    Hand.Color = NColor.BackColor
    CHand.Color = NColor.BackColor
    FHS NColor.BackColor
End Sub

Private Sub TColor_Click()
    If TopMostMode Then SetWindowPos MoreInfo.hWnd, HWND_TOPMOST, _
        0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    MoreInfo.Show 1
End Sub

Public Sub GetColorInfo(ByVal CL As Long)
    Static C As New ColorInfo, K As String
       
    If CL < 0 Then Exit Sub
    TColor.BackColor = CL: C.Color = CL
    
    Info(0).Tag = C.getStaName
    Info(0).Caption = "�зǦW�G" & Info(0).Tag
    Info(0).SetToolTipText IIf(Info(0).Tag = "�L", "�S���i�Ϊ���", "���@�U�H�ƻs��ŶKï")
    
    Info(1).Tag = C.getExtName
    LB.Caption = "�����W�G" & Info(1).Tag
    If LB.Width > Info(2).Width / Screen.TwipsPerPixelX - 5 Then LB.Caption = LB.Caption & "..."
    Do While LB.Width > Info(2).Width / Screen.TwipsPerPixelX - 5
        LB.Caption = Mid(LB.Caption, 1, Len(LB.Caption) - 4) & "..."
    Loop
    
    Info(1).Caption = LB.Caption
    Info(1).SetToolTipText IIf(Info(1).Tag = "�L", "�S���i�Ϊ���", "���@�U�H�ƻs�u" & Info(1).Tag & "�v��ŶKï")
    
    Info(2).Tag = IIf(LeadingSharp, "#", "") & C.getRRGGBB
    Info(2).Caption = "�зǭȡG" & Info(2).Tag
    
    Info(3).Tag = IIf(C.getRGB = "�L", "", IIf(LeadingSharp, "#", "")) & C.getRGB
    Info(3).Caption = "²�g�ȡG" & Info(3).Tag
    Info(3).SetToolTipText IIf(Info(3).Tag = "�L", "�S���i�Ϊ���", "���@�U�H�ƻs��ŶKï")
    
    Info(4).Tag = IIf(LeadingSharp, "#", "") & C.getNRGB
    Info(4).Caption = "����ȡG" & Info(4).Tag
    
    Info(5).Tag = IIf(LeadingSharp, "#", "") & C.getSRGB
    Info(5).Caption = "�w���ȡG" & Info(5).Tag
    
    NColor.BackColor = C.getNColor
    SColor.BackColor = C.getSColor
End Sub

Private Sub Info_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Info(Index).ForeColor = vbBlue
End Sub

Private Sub Info_MouseOut(Index As Integer)
    Info(Index).ForeColor = vbBlack
End Sub

Private Sub Info_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static txt As String, C As New ColorInfo
    If Button = 1 Then
        C.Color = TColor.BackColor
        txt = Info(Index).Tag
        If txt <> "�L" Then
            Clipboard.Clear
            Clipboard.SetText txt
        End If
        Info_MouseOut Index
    End If
End Sub








' ---------------------------------
' ��ʰt��P�L��t��
' ---------------------------------

Private Sub Hand_Change()
    GetColorInfo Hand.Color
End Sub

Private Sub CHand_Change()
    GetColorInfo CHand.Color
End Sub

Private Sub hiddenText1_GotFocus()
    ' �o�Өƥ�P�U�@�Өƥ�O���F�d�I�q Hand �]�X�Ӫ� TabStop �Ϊ�
    ' ������ VB ������ TabStop ����A�b��V UserControl �ɷ|�X���D
    ' �]���ĥΦ����j�@�k
    If Fra(0).Visible Then Hand.SetTextFocus 1
    If Fra(1).Visible Then CHand.SetTextFocus 1
End Sub

Private Sub hiddenText2_GotFocus()
    ' ���~�����p�ߡA�� Fra(0) �Q���îɳo��Ӫ���|��o�n�I�A
    ' ���Y�A���Ұʻy�k�չϥh���Q���ê�������o�n�I�N�|�o���Y�����~�A
    ' �]���u��� Fra(0) �S�Q���îɤ~����y�k
    If Fra(0).Visible Then Hand.SetTextFocus 3
    If Fra(1).Visible Then CHand.SetTextFocus 4
End Sub









' ---------------------------------
' ��t��L
' ---------------------------------

Private Sub HSColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then HSColor_MouseMove Button, Shift, X, Y
End Sub

Private Sub HSColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static C As Object
    If Button = 1 Then
        If HSVMode Then Set C = New ColorHSV Else Set C = New ColorHSL
        C.Color = HSColor.Color: C.I = IColor.I
        GetColorInfo C.Color
        IColor.H = C.H: IColor.S = C.S: IColor.Draw
    End If
End Sub

Private Sub IColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then GetColorInfo IColor.Color
End Sub

Private Sub IColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then GetColorInfo IColor.Color
End Sub

Private Sub FHS(CC As Long)
    HSColor.Color = CC: IColor.Color = CC
    GetColorInfo CC
    IColor.Draw
End Sub

Private Sub isHSV_Click()
    HSVMode = True: HSColor.Draw: FHS TColor.BackColor
End Sub

Private Sub isHSL_Click()
    HSVMode = False: HSColor.Draw: FHS TColor.BackColor
End Sub

Private Sub isHSV_GotFocus()
    KBDOptionFocused = 3
End Sub

Private Sub isHSL_GotFocus()
    KBDOptionFocused = 4
End Sub

Private Sub isHSV_LostFocus()
    KBDOptionFocused = 0
End Sub

Private Sub isHSL_LostFocus()
    KBDOptionFocused = 0
End Sub






' ---------------------------------
' �ù��ߦ�
' ---------------------------------

Private Sub picScr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Not ScrCapturing Then GetColorInfo picScr.Point(X, Y)
End Sub

Private Sub picScr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Not ScrCapturing Then GetColorInfo picScr.Point(X, Y)
End Sub

Private Sub ScrStart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ScrCapturing Then
        ScrCaptureEnd
    Else
        TempColor = TColor.BackColor
        Me.Hide
        TimerScr.Enabled = True
    End If
End Sub

Public Sub ScrCaptureEnd(Optional ByVal Cancel As Boolean = True)
    ScrStart.Caption = "�}�l"
    ScrStart.SetToolTipText "���@�U�H�}�l�^���ù���m"
    picScr.ToolTipText = "���@�U�H��ܦ�m"
    picScr.MousePointer = 10
    M_Fun.Enabled = True: M_Pic.Enabled = True: M_Opt.Enabled = True
    Unload FScreen
    If Not TopMostMode Then SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    If Cancel Then GetColorInfo TempColor
    ScrCapturing = False
End Sub

Private Sub TimerScr_Timer()
    TimerScr.Enabled = False
    FScreen.P.Width = Screen.Width / Screen.TwipsPerPixelX
    FScreen.P.Height = Screen.Height / Screen.TwipsPerPixelY
    BitBlt FScreen.P.hDC, 0, 0, FScreen.P.Width, FScreen.P.Height, GetDC(0), 0, 0, vbSrcCopy
    M_Fun.Enabled = False: M_Pic.Enabled = False: M_Opt.Enabled = False
    SetWindowPos FScreen.hWnd, HWND_TOPMOST, 0, 0, FScreen.P.Width, FScreen.P.Height, SWP_SHOWWINDOW
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    ScrStart.Caption = "����"
    ScrStart.SetToolTipText "���@�U�H�����^���ù���m"
    picScr.ToolTipText = "�ù�������j��"
    ScrCapturing = True
End Sub









' ---------------------------------
' ������L�P�t�ΦW��
' ---------------------------------

Private Sub PaintWebColor()
    Static ScX As Integer, ScY As Integer, bHeight As Integer
    Static I As Integer, j As Integer, K As Integer, SC As Long, T As Integer
        
    ScX = Screen.TwipsPerPixelX: ScY = Screen.TwipsPerPixelY
    bHeight = PK.Height / 8: KP.Height = bHeight * 50 + ScY
    PK.Height = bHeight * 8 + ScY: KVS.Height = PK.Height: SVS.Height = PK.Height
    KL(0).Top = 4 * ScY: KL(0).Left = 4 * ScX
    SK(0).Top = bHeight: SK(0).Height = bHeight * 2 + ScY
    For I = 0 To 15
        SC = dColor(I + 1, 1)
        SK(0).Line ((I Mod 8) * 255, Int(I / 8) * bHeight)-((I Mod 8 + 1) * 255, Int(I / 8 + 1) * bHeight), SC, BF
        SK(0).Line ((I Mod 8) * 255, Int(I / 8) * bHeight)-((I Mod 8 + 1) * 255, Int(I / 8 + 1) * bHeight), vbBlack, B
    Next I
    KL(1).Top = bHeight * 3 + 4 * ScY: KL(1).Left = 4 * ScX
    SK(1).Top = bHeight * 4: SK(1).Height = bHeight * 18 + ScY
    For I = 0 To 139
        SC = YColor(I + 1, 1)
        SK(1).Line ((I Mod 8) * 255, Int(I / 8) * bHeight)-((I Mod 8 + 1) * 255, Int(I / 8 + 1) * bHeight), SC, BF
        SK(1).Line ((I Mod 8) * 255, Int(I / 8) * bHeight)-((I Mod 8 + 1) * 255, Int(I / 8 + 1) * bHeight), vbBlack, B
    Next I
    SK(1).Line (0, 0)-(SK(1).ScaleWidth - ScX, SK(1).ScaleHeight - ScY), vbBlack, B
    KL(2).Top = bHeight * 22 + 4 * ScY: KL(2).Left = 4 * ScX
    SK(2).Top = bHeight * 23: SK(2).Height = bHeight * 27 + ScY
    For I = 0 To 15 Step 3
        For j = 0 To 15 Step 3
            For K = 0 To 15 Step 3
                SC = K * &H110000 + j * &H1100& + I * &H11: T = I * 12 + j * 2 + K / 3
                SK(2).Line ((T Mod 8) * 255, Int(T / 8) * bHeight)-((T Mod 8 + 1) * 255, Int(T / 8 + 1) * bHeight), SC, BF
                SK(2).Line ((T Mod 8) * 255, Int(T / 8) * bHeight)-((T Mod 8 + 1) * 255, Int(T / 8 + 1) * bHeight), vbBlack, B
            Next K
        Next j
    Next I
    KP.Line (0, 0)-(KP.ScaleWidth - ScX, KP.ScaleHeight - ScY), vbBlack, B
    KVS.Max = (KP.Height - PK.Height)
    KVS.SmallChange = bHeight
    KVS.LargeChange = bHeight * 8
End Sub

Private Sub SK_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static P(4) As Long, I As Integer
    If Button = 1 And Y + SK(Index).Top > KVS.Value _
        And Y + SK(Index).Top < KVS.Value + KVS.LargeChange Then
    
        ' �T�w�ƹ����b���j�u�W
        P(1) = SK(Index).Point(X - Screen.TwipsPerPixelX, Y - Screen.TwipsPerPixelX)
        P(2) = SK(Index).Point(X + Screen.TwipsPerPixelX, Y - Screen.TwipsPerPixelX)
        P(3) = SK(Index).Point(X - Screen.TwipsPerPixelX, Y + Screen.TwipsPerPixelX)
        P(4) = SK(Index).Point(X + Screen.TwipsPerPixelX, Y + Screen.TwipsPerPixelX)
        If Not SK(Index).Point(X, Y) = vbBlack Or _
            P(1) = vbBlack And P(2) = vbBlack And P(3) = vbBlack And P(4) = vbBlack Then _
            GetColorInfo SK(Index).Point(X, Y)
    
        ' �ǻ��ƥ��t�~��� SK �W
        If Not EventRaised Then
            EventRaised = True
            For I = 0 To 2
                If Index <> I Then SK_MouseMove I, Button, Shift, X, Y + SK(Index).Top - SK(I).Top
            Next I
            EventRaised = False
        End If
    End If
End Sub

Private Sub SK_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SK_MouseMove Index, Button, Shift, X, Y
End Sub

Private Sub Sy_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clipboard.Clear
    Clipboard.SetText Sys(Index)
End Sub

Private Sub SVS_Change()
    SVS.Value = CInt(SVS.Value / SVS.SmallChange) * SVS.SmallChange
    SP.Top = -SVS.Value
End Sub

Private Sub SVS_Scroll()
    SVS_Change
End Sub

Private Sub KVS_Change()
    KVS.Value = CInt(KVS.Value / KVS.SmallChange) * KVS.SmallChange
    KP.Top = -KVS.Value
End Sub

Private Sub KVS_Scroll()
    KVS_Change
End Sub











' ---------------------------------
' �Ϥ��ߦ�
' ---------------------------------

Private Sub M_Pic_Click()
    Static K As Object
    For Each K In Fra
        If K.Index = 6 Then K.Visible = True Else K.Visible = False
    Next K
End Sub

Private Sub MP_Cli_Click()
    Static Files() As String
          
    If cImage.LoadPicture_ClipBoard() Then
        LoadcImage
    ElseIf cImage.GetPastedFileNames(Files()) > 0 Then
        If TryLoadPicture(Files(1)) Then LoadcImage _
            Else MsgBox "�K�W���ɮפ��O�䴩���Ϥ��ɡC", vbCritical
    Else
        MsgBox "�Х��N�Ϥ��ƻs��ŶKï���C", vbCritical
    End If
    Exit Sub
End Sub

Private Sub MP_Fil_Click()
    Static C As New CDialogW, sFile As String, sFileTitle As String
    
    If C.VBGetOpenFileName(sFile, sFileTitle, , , , , _
        "�Ҧ��䴩���榡|*.bmp;*.jpg;*.jpeg;*.gif;*.png;*.ico;*.cur;*.wmf;*.emf;*.tga;*.tiff|" & _
        "�Ϥ��� (*.bmp;*.jpg;*.jpeg;*.gif;*.png;*.tga;*.tiff)|*.bmp;*.jpg;*.jpeg;*.gif;*.png;*.tga;*.tiff|" & _
        "�ϥ��� (*.ico;*.cur)|*.ico;*.cur|" & _
        "���~�� (*.wmf;*.emf)|*.wmf;*.emf|" & _
        "�Ҧ��ɮ� (*.*)|*.*", , , _
        "�}�ҹϤ��ɮ�", "TXT", Me.hWnd, OFN_HIDEREADONLY) Then
        
        If TryLoadPicture(sFile) Then LoadcImage _
            Else MsgBox "�ɮ׶}�ҵo�Ϳ��~�C", vbCritical
    End If
End Sub

Private Function TryLoadPicture(ByVal FileName As String) As Boolean
    Static SavedPointer As Integer
    SavedPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    TryLoadPicture = cImage.LoadPicture_File(FileName, 256, 256)
    Screen.MousePointer = SavedPointer
End Function

Private Sub LoadcImage()
    Static I As Integer, SavedPointer As Integer, B As Long, OB As Long
      
    SavedPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    ' ���o�t�Ϊ��C��
    B = GetSysColorBrush(COLOR_BTNFACE)
    
    ' �إ߭��
    PWidth = cImage.Width: PHeight = cImage.Height
    DPBitmap(XXCIndex) = CreateDIBSec(Me.hDC, PWidth, PHeight)
    OB = SelectObject(DPDC(XXCIndex), DPBitmap(XXCIndex))
    DeleteObject OB
    FillRect DPDC(XXCIndex), CreateRect(0, 0, PWidth, PHeight), B
    cImage.Render DPDC(XXCIndex), 0, 0, PWidth, PHeight
         
    XXMaxIndex = MS_x.Count
    For I = XXMaxIndex To 1 Step -1
        If MS_x(I).Tag = 1 Then Exit For
    
        ' �ھڹϤ��j�p�]�w���v����
        If PWidth * MS_x(I).Tag < PP.ScaleWidth And PHeight * MS_x(I).Tag < PP.ScaleWidth Then
            MS_x(I).Visible = False: XXMaxIndex = I - 1
        Else
            MS_x(I).Visible = True
            
            ' �إ��Y�ϰ}�C
            DPBitmap(I) = CreateDIBSec(Me.hDC, PWidth * MS_x(I).Tag, PHeight * MS_x(I).Tag)
            OB = SelectObject(DPDC(I), DPBitmap(I))
            DeleteObject OB
            FillRect DPDC(I), CreateRect(0, 0, PWidth * MS_x(I).Tag, Height * MS_x(I).Tag), B
            cImage.Render DPDC(I), 0, 0, PWidth * MS_x(I).Tag, PHeight * MS_x(I).Tag
        End If
    Next I
    
    MS_x_Click XXCIndex             ' �N���v�٭�
    Hs.Value = 0: Vs.Value = 0      ' �N���b�k��
    OldX = 0: OldY = 0
    
    PictureLoaded = True
    PicSize
    M_Pic_Click
    
    Screen.MousePointer = SavedPointer
End Sub

Private Sub PicSize()
    PPWidth = CInt(PP.ScaleWidth / XX): PPHeight = CInt(PP.ScaleHeight / XX)
    
    PaintLock = True
    If PWidth > PPWidth Then
        Hs.Enabled = True
        Hs.Max = PWidth - PPWidth
        Hs.LargeChange = 3 / 4 * PPWidth
        Hs.SmallChange = IIf(XX < 1, 1 / XX, 1)
        Hs.Value = ValueInRange(OldX - Int((PPWidth + 1) / 2), 0, Hs.Max)
    Else
        Hs.Enabled = False
        Hs.Value = 0
    End If
    If PHeight > PPHeight Then
        Vs.Enabled = True
        Vs.Max = PHeight - PPHeight
        Vs.LargeChange = 3 / 4 * PPHeight
        Vs.SmallChange = IIf(XX < 1, 1 / XX, 1)
        Vs.Value = ValueInRange(OldY - Int((PPHeight + 1) / 2), 0, Vs.Max)
    Else
        Vs.Enabled = False
        Vs.Value = 0
    End If
    
    PP.Cls
    PaintLock = False
    
    PictureScroll
End Sub

Private Sub Hs_Change()
    PictureScroll
End Sub

Private Sub Hs_Scroll()
    PictureScroll
End Sub

Private Sub Vs_Change()
    PictureScroll
End Sub

Private Sub Vs_Scroll()
    PictureScroll
End Sub

Private Sub PictureScroll()
    If Not PaintLock And PictureLoaded Then
        If XX >= 1 Then
            StretchBlt PP.hDC, 0, 0, PP.ScaleWidth, PP.ScaleHeight, DPDC(XXCIndex), _
                Hs.Value, Vs.Value, PPWidth, PPHeight, vbSrcCopy
        Else
            StretchBlt PP.hDC, 0, 0, PP.ScaleWidth, PP.ScaleHeight, DPDC(XXIndex), _
                Hs.Value * XX, Vs.Value * XX, PP.ScaleWidth, PP.ScaleHeight, vbSrcCopy
        End If
        OldX = Hs.Value + Int((PPWidth + 1) / 2)
        OldY = Vs.Value + Int((PPHeight + 1) / 2)
    End If
End Sub

Private Sub PP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' �p�G�w�g������Q���U�A�N���z�|�s�����U
    If PictureMouseButton <> 0 Then Exit Sub
    
    If PictureLoaded And Button = 1 And PP.MousePointer = 10 Then
        LockCursor PP.hWnd, 2
        GetColorInfo PP.Point(X, Y)
    End If
    If PictureLoaded And (Button = 2 Or (Button = 1 And PP.MousePointer = 15)) Then
        ' �Ȯ������t�ο��A�K�o�즲�ɻ~Ĳ
        If Button = 2 Then EnableSystemMenu Me.hWnd, False
        PP.MousePointer = 15
        PictureX = X + Hs.Value * XX: PictureY = Y + Vs.Value * XX
    End If
    PictureMouseButton = Button
End Sub

Private Sub PP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PictureLoaded And PictureMouseButton = 1 And PP.MousePointer = 10 Then GetColorInfo PP.Point(X, Y)
    If PictureLoaded And (PictureMouseButton = 2 Or (PictureMouseButton = 1 And PP.MousePointer = 15)) Then
        If Hs.Enabled Then Hs.Value = ValueInRange(CInt((PictureX - X) / XX), 0, Hs.Max)
        If Vs.Enabled Then Vs.Value = ValueInRange(CInt((PictureY - Y) / XX), 0, Vs.Max)
    End If
    ControlX = X: ControlY = Y
End Sub

Public Sub PP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PictureMouseButton = 1 And PictureLoaded Then
        UnLockCursor
        GetColorInfo PP.Point(X, Y)
    End If
    ' �u�����}�����䵥�󤧫e���U������ɤ~��@�O�u����}
    If PictureMouseButton = Button Then PictureMouseButton = 0
    EnableSystemMenu Me.hWnd, True
End Sub

Private Sub PP_Paint()
    If PictureLoaded Then PictureScroll
End Sub

Private Sub Siz_Click()
    Me.PopupMenu M_Siz
End Sub

Public Sub MS_x_Click(Index As Integer)
    If PictureLoaded Then
        MS_x(XXIndex).Checked = False
        MS_x(Index).Checked = True
        XX = MS_x(Index).Tag
        XXIndex = Index
        PicSize
        ' �p�G���b�즲�Ϥ��A�ץ��즲�Ѧ��I
        If PictureMouseButton = 2 Or (PictureMouseButton = 1 And PP.MousePointer = 15) Then
            PictureX = ControlX + Hs.Value * XX: PictureY = ControlY + Vs.Value * XX
        End If
    End If
End Sub



















' ---------------------------------
' ���[����
' ---------------------------------

' �s�����B�z��
Private Sub Timer_Timer()
    Static Key As Boolean, Speed As Integer
    Static C As Object
    
    ' ���K�B�z PP �����ܤ�r
    If PPToolTipTimeout > 0 And PPToolTipTimeout < 100 Then
        PPToolTipTimeout = PPToolTipTimeout + 1
    ElseIf PPToolTipTimeout = 100 Then
        PP.ToolTipText = PP.Tag
        PPToolTipTimeout = 0
    End If
    
    If PictureLoaded And Fra(6).Visible Then
        If PictureMouseButton = 0 Then PP.MousePointer = IIf(GetKeyState(vbKeySpace) And &H1000, 15, 10)
        If PP.MousePointer = 10 Then
            Key = False: Speed = 1.5 ^ (XXIndex - 1)
            If (GetKeyState(vbKeyUp) And &H1000) Then ScrollBarScroll Vs, -Speed: Key = True
            If (GetKeyState(vbKeyDown) And &H1000) Then ScrollBarScroll Vs, Speed: Key = True
            If (GetKeyState(vbKeyLeft) And &H1000) Then ScrollBarScroll Hs, -Speed: Key = True
            If (GetKeyState(vbKeyRight) And &H1000) Then ScrollBarScroll Hs, Speed: Key = True
            If Key Then RaiseMouseMove
        End If
    ElseIf Fra(2).Visible Then
        Key = False
        If (GetKeyState(vbKeyUp) And &H1000) Then HSColor.PointerMove 0, -1: Key = True
        If (GetKeyState(vbKeyDown) And &H1000) Then HSColor.PointerMove 0, 1: Key = True
        If (GetKeyState(vbKeyLeft) And &H1000) Then HSColor.PointerMove -1, 0: Key = True
        If (GetKeyState(vbKeyRight) And &H1000) Then HSColor.PointerMove 1, 0: Key = True
        If Key Then
            If HSVMode Then Set C = New ColorHSV Else Set C = New ColorHSL
            C.Color = HSColor.Color: C.I = IColor.I
            GetColorInfo C.Color: IColor.H = C.H
            IColor.S = C.S: IColor.Draw
        End If
    End If
End Sub

' �����ֳt��B�z��
Public Sub Form_GlobalKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    Static S As String

    ' �K�W Ctrl-V
    If KeyCode = vbKeyV And (Shift And vbCtrlMask) Then
        If Clipboard.GetFormat(1) Then
            S = Clipboard.GetText
            If Not KBDHooked Or Len(S) > 3 Or _
                (DecMode And Not IsNumeric(S)) Or _
                (Not DecMode And Not IsNumeric("&H" & S)) _
                Then MF_Pas_Click
        Else
            MP_Cli_Click
        End If
    End If
    
    ' �ƻs Ctrl-C
    If Not KBDHooked And KeyCode = vbKeyC And (Shift And vbCtrlMask) Then
        Info_MouseUp 2, 1, 0, 0, 0
    End If
    
    ' �W�U������
    If Not Fra(6).Visible Then
        If KeyCode = vbKeyPageUp And (Shift And vbCtrlMask) Then MF_Fra_Click ValueInRange(origFrame - 1, 0, 5)
        If KeyCode = vbKeyPageDown And (Shift And vbCtrlMask) Then MF_Fra_Click ValueInRange(origFrame + 1, 0, 5)
    End If
     
    ' ���b����
    If Fra(5).Visible And (Shift And vbCtrlMask) = 0 Then
        If KeyCode = vbKeyUp Then ScrollBarScroll SVS, -SVS.SmallChange
        If KeyCode = vbKeyDown Then ScrollBarScroll SVS, SVS.SmallChange
        If KeyCode = vbKeyPageUp Then ScrollBarScroll SVS, -SVS.LargeChange
        If KeyCode = vbKeyPageDown Then ScrollBarScroll SVS, SVS.LargeChange
        If KeyCode = vbKeyHome Then SVS.Value = 0
        If KeyCode = vbKeyEnd Then SVS.Value = SVS.Max
    ElseIf Fra(4).Visible And (Shift And vbCtrlMask) = 0 Then
        If KeyCode = vbKeyUp Then ScrollBarScroll KVS, -KVS.SmallChange
        If KeyCode = vbKeyDown Then ScrollBarScroll KVS, KVS.SmallChange
        If KeyCode = vbKeyPageUp Then ScrollBarScroll KVS, -KVS.LargeChange
        If KeyCode = vbKeyPageDown Then ScrollBarScroll KVS, KVS.LargeChange
        If KeyCode = vbKeyHome Then KVS.Value = 0
        If KeyCode = vbKeyEnd Then KVS.Value = KVS.Max
    End If
End Sub

' �u���B�z��
Public Sub Form_MouseWheel(ByVal wParam As Long)
    Static OX As Integer
    
    If Fra(0).Visible Then
        Hand.Scroll IIf(wParam > 0, 1, -1) * 5
    ElseIf Fra(1).Visible Then
        CHand.Scroll IIf(wParam > 0, 1, -1) * 5
    ElseIf Fra(2).Visible Then
        IColor.I = IColor.I + IIf(wParam > 0, 1, -1) * 5
        GetColorInfo IColor.Color
    ElseIf Fra(4).Visible Then
        ScrollBarScroll KVS, IIf(wParam > 0, -1, 1) * KVS.SmallChange
    ElseIf Fra(5).Visible Then
        ScrollBarScroll SVS, IIf(wParam > 0, -1, 1) * SVS.SmallChange
    ElseIf Fra(6).Visible Then
        OX = XXIndex
        If wParam > 0 Then OX = OX - 1 Else OX = OX + 1
        If OX >= 1 And OX <= XXMaxIndex Then MS_x_Click OX
        PP.ToolTipText = MS_x(XXIndex).Caption
        PPToolTipTimeout = 1
    End If
End Sub

' �ɮש��B�z��
Public Sub Form_FileDragDrop(ByVal FileName As String)
    If TryLoadPicture(FileName) Then LoadcImage _
        Else MsgBox "�o���ɮפ��O�䴩���Ϥ��ɡC", vbCritical
End Sub


