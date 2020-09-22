VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.1#0"; "osenxpsuite.ocx"
Begin VB.Form Frm_addressbook 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Address Book"
   ClientHeight    =   8790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_addressbook.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   586
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   550
   StartUpPosition =   2  'CenterScreen
   Begin osenxpsuite.OsenXPFrame OsenXPFrame2 
      Height          =   1725
      Left            =   180
      TabIndex        =   70
      Top             =   6180
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   3043
      Caption         =   "User Action:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   12570832
      Begin osenxpsuite.OsenXPButton CmdExit 
         Height          =   345
         Left            =   1380
         TabIndex        =   27
         Top             =   1230
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "&Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_addressbook.frx":058A
         PICN            =   "frm_addressbook.frx":05A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         XPBlendPicture  =   0   'False
         GradientColor1  =   16772057
         GradientColor2  =   16777215
      End
      Begin osenxpsuite.OsenXPButton CmdRefresh 
         Height          =   345
         Left            =   150
         TabIndex        =   28
         Top             =   1230
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "&Refresh"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_addressbook.frx":0940
         PICN            =   "frm_addressbook.frx":095C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         XPBlendPicture  =   0   'False
         GradientColor1  =   16772057
         GradientColor2  =   16777215
      End
      Begin osenxpsuite.OsenXPButton CmdDelete 
         Height          =   345
         Left            =   1380
         TabIndex        =   26
         Top             =   780
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "&Delete"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_addressbook.frx":0CF6
         PICN            =   "frm_addressbook.frx":0D12
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         XPBlendPicture  =   0   'False
         GradientColor1  =   16772057
         GradientColor2  =   16777215
      End
      Begin osenxpsuite.OsenXPButton CmdSearch 
         Height          =   345
         Left            =   150
         TabIndex        =   25
         Top             =   780
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "&Search"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_addressbook.frx":12AC
         PICN            =   "frm_addressbook.frx":12C8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         XPBlendPicture  =   0   'False
         GradientColor1  =   16772057
         GradientColor2  =   16777215
      End
      Begin osenxpsuite.OsenXPButton CmdAction 
         Height          =   345
         Index           =   1
         Left            =   1380
         TabIndex        =   24
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "&Edit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_addressbook.frx":1662
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         OffsetLeft      =   0
         OffsetTop       =   0
         XPBlendPicture  =   0   'False
         GradientColor1  =   16772057
         GradientColor2  =   16777215
      End
      Begin osenxpsuite.OsenXPButton CmdAction 
         Height          =   345
         Index           =   0
         Left            =   150
         TabIndex        =   23
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "&Add New"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_addressbook.frx":167E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         OffsetLeft      =   0
         OffsetTop       =   0
         XPBlendPicture  =   0   'False
         GradientColor1  =   16772057
         GradientColor2  =   16777215
      End
   End
   Begin VB.ComboBox CboScheme 
      Height          =   315
      ItemData        =   "frm_addressbook.frx":169A
      Left            =   1440
      List            =   "frm_addressbook.frx":16A7
      TabIndex        =   68
      Text            =   "XPBlue"
      Top             =   5730
      Width           =   1365
   End
   Begin osenxpsuite.OsenXPStatusBar OsenXPStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   46
      Top             =   8385
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   714
      BackColor       =   14936810
      ForeColor       =   -2147483630
      ForeColorDissabled=   -2147483631
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   1
      PWidth1         =   1000
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "Address Book"
      pTextAlignment1 =   0
      PanelPicture1   =   "frm_addressbook.frx":16C6
      PanelPicAlignment1=   0
      pBckgColor1     =   0
      pGradient1      =   0
      pEdgeSpacing1   =   0
      pEdgeInner1     =   0
      pEdgeOuter1     =   0
      DrawMode        =   1
      HaveXPForm      =   -1  'True
   End
   Begin osenxpsuite.OsenXPTab OsenXPTab1 
      Height          =   3465
      Left            =   2970
      TabIndex        =   33
      Top             =   4440
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6112
      TabHeight       =   22
      BackColor       =   16777215
      ForeColor       =   -2147483630
      ForeColorActive =   9982008
      ForeColorHot    =   16711680
      FrameColor      =   12164479
      MaskColor       =   16711935
      SelectedTab     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumberOfTabs    =   2
      TabWidth1       =   57
      TabText1        =   "Home"
      TabEnabled1     =   -1  'True
      TabPicture1     =   "frm_addressbook.frx":16E2
      TabWidth2       =   58
      TabText2        =   "Work"
      TabEnabled2     =   -1  'True
      TabPicture2     =   "frm_addressbook.frx":1A34
      BackColorParent =   14215660
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3075
         Index           =   2
         Left            =   30
         ScaleHeight     =   3075
         ScaleWidth      =   4965
         TabIndex        =   50
         Top             =   5550
         Width           =   4965
         Begin VB.ComboBox CboCountry 
            DataField       =   "wcountry"
            Height          =   315
            Index           =   0
            Left            =   960
            TabIndex        =   20
            Text            =   "Combo1"
            Top             =   2310
            Width           =   3855
         End
         Begin osenxpsuite.OsenXPTextBox OxtData 
            Height          =   645
            Index           =   7
            Left            =   960
            TabIndex        =   16
            Top             =   840
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   1138
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            DataField       =   "wstreetaddress"
            MultiLine       =   -1  'True
         End
         Begin osenxpsuite.OsenXPTextBox OxtData 
            Height          =   285
            Index           =   8
            Left            =   960
            TabIndex        =   17
            Top             =   1590
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   503
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            DataField       =   "wcity"
         End
         Begin osenxpsuite.OsenXPTextBox OxtData 
            Height          =   285
            Index           =   9
            Left            =   960
            TabIndex        =   18
            Top             =   1950
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   503
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            DataField       =   "wstate"
         End
         Begin osenxpsuite.OsenXPTextBox OxtData 
            Height          =   285
            Index           =   10
            Left            =   3420
            TabIndex        =   19
            Top             =   1950
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            DataField       =   "wzipcode"
         End
         Begin osenxpsuite.OsenXPTextBox OxtData 
            Height          =   285
            Index           =   11
            Left            =   960
            TabIndex        =   21
            Top             =   2700
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   503
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            DataField       =   "wphone"
         End
         Begin osenxpsuite.OsenXPTextBox OxtData 
            Height          =   285
            Index           =   12
            Left            =   3060
            TabIndex        =   22
            Top             =   2700
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            DataField       =   "wfax"
         End
         Begin osenxpsuite.OsenXPTextBox OxtData 
            Height          =   285
            Index           =   19
            Left            =   1410
            TabIndex        =   14
            Top             =   90
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   503
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            DataField       =   "companyname"
         End
         Begin osenxpsuite.OsenXPTextBox OxtData 
            Height          =   285
            Index           =   20
            Left            =   960
            TabIndex        =   15
            Top             =   450
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   503
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            DataField       =   "jobtitle"
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Job Title:"
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   67
            Top             =   480
            Width           =   660
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company Name:"
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   66
            Top             =   120
            Width           =   1185
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Country:"
            Height          =   195
            Index           =   13
            Left            =   180
            TabIndex        =   57
            Top             =   2340
            Width           =   645
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
            Height          =   195
            Index           =   12
            Left            =   2700
            TabIndex        =   56
            Top             =   2730
            Width           =   330
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone:"
            Height          =   195
            Index           =   11
            Left            =   180
            TabIndex        =   55
            Top             =   2730
            Width           =   510
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Zip Code:"
            Height          =   195
            Index           =   10
            Left            =   2610
            TabIndex        =   54
            Top             =   1980
            Width           =   690
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "State:"
            Height          =   195
            Index           =   9
            Left            =   150
            TabIndex        =   53
            Top             =   1980
            Width           =   450
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "City:"
            Height          =   195
            Index           =   8
            Left            =   150
            TabIndex        =   52
            Top             =   1620
            Width           =   345
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Street:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   51
            Top             =   810
            Width           =   510
         End
      End
      Begin VB.PictureBox PicPage 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3075
         Index           =   1
         Left            =   30
         ScaleHeight     =   3075
         ScaleWidth      =   4995
         TabIndex        =   58
         Top             =   360
         Width           =   4995
         Begin VB.ComboBox CboCountry 
            DataField       =   "hcountry"
            Height          =   315
            Index           =   1
            Left            =   960
            TabIndex        =   11
            Text            =   "Combo1"
            Top             =   1890
            Width           =   3825
         End
         Begin osenxpsuite.OsenXPTextBox OxtData 
            Height          =   645
            Index           =   13
            Left            =   960
            TabIndex        =   7
            Top             =   270
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   1138
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            DataField       =   "hstreetaddress"
            MultiLine       =   -1  'True
         End
         Begin osenxpsuite.OsenXPTextBox OxtData 
            Height          =   285
            Index           =   14
            Left            =   960
            TabIndex        =   8
            Top             =   1050
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   503
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            DataField       =   "hcity"
         End
         Begin osenxpsuite.OsenXPTextBox OxtData 
            Height          =   285
            Index           =   15
            Left            =   960
            TabIndex        =   9
            Top             =   1470
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   503
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            DataField       =   "hstate"
         End
         Begin osenxpsuite.OsenXPTextBox OxtData 
            Height          =   285
            Index           =   16
            Left            =   3420
            TabIndex        =   10
            Top             =   1470
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            DataField       =   "hzipcode"
         End
         Begin osenxpsuite.OsenXPTextBox OxtData 
            Height          =   285
            Index           =   17
            Left            =   960
            TabIndex        =   12
            Top             =   2340
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   503
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            DataField       =   "hphone"
         End
         Begin osenxpsuite.OsenXPTextBox OxtData 
            Height          =   285
            Index           =   18
            Left            =   3060
            TabIndex        =   13
            Top             =   2340
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            DataField       =   "hfax"
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Street:"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   510
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "City:"
            Height          =   195
            Index           =   19
            Left            =   150
            TabIndex        =   64
            Top             =   1080
            Width           =   345
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "State:"
            Height          =   195
            Index           =   18
            Left            =   150
            TabIndex        =   63
            Top             =   1500
            Width           =   450
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Zip Code:"
            Height          =   195
            Index           =   17
            Left            =   2610
            TabIndex        =   62
            Top             =   1500
            Width           =   690
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone:"
            Height          =   195
            Index           =   16
            Left            =   180
            TabIndex        =   61
            Top             =   2370
            Width           =   510
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
            Height          =   195
            Index           =   15
            Left            =   2700
            TabIndex        =   60
            Top             =   2370
            Width           =   330
         End
         Begin VB.Label LblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Country:"
            Height          =   195
            Index           =   14
            Left            =   180
            TabIndex        =   59
            Top             =   1920
            Width           =   645
         End
      End
   End
   Begin osenxpsuite.OsenXPListBox LstContact 
      Height          =   3945
      Left            =   210
      TabIndex        =   32
      Top             =   1620
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   6959
      Appearance      =   0
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSelected    =   16576
      BackSelected    =   10841658
      BackSelectedG1  =   16777215
      BackSelectedG2  =   14854529
      ItemHeight      =   20
      ItemHeightAuto  =   0   'False
      ItemOffset      =   2
      SelectModeStyle =   2
      ShowHeader      =   -1  'True
      Columns         =   1
      CT1             =   "Contact Name"
      CA1             =   0
      CW1             =   150
   End
   Begin osenxpsuite.MyADODC MyData 
      Height          =   375
      Left            =   150
      TabIndex        =   31
      Top             =   7950
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      BeginProperty FontButton {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "MyADODC1"
   End
   Begin VB.PictureBox PicHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   0
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   550
      TabIndex        =   30
      Top             =   450
      Width           =   8250
      Begin VB.Line Line1 
         X1              =   0
         X2              =   550
         Y1              =   66
         Y2              =   66
      End
      Begin VB.Label lbdescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Description of address book database, put here ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   780
         TabIndex        =   43
         Top             =   510
         Width           =   4785
      End
      Begin VB.Label Lbtitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address Book Database"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   42
         Top             =   180
         Width           =   2325
      End
      Begin VB.Image ImgLogo 
         Height          =   600
         Left            =   7350
         Picture         =   "frm_addressbook.frx":1D86
         Stretch         =   -1  'True
         Top             =   180
         Width           =   630
      End
   End
   Begin osenxpsuite.OsenXPForm XPF 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Address Book"
      TitleTop        =   7
      icon            =   "frm_addressbook.frx":2650
      MaximizeEnabled =   0   'False
   End
   Begin osenxpsuite.OsenXPFrame OsenXPFrame1 
      Height          =   2865
      Left            =   2970
      TabIndex        =   34
      Top             =   1500
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5054
      Caption         =   "Personal Info:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   12570832
      Image           =   "frm_addressbook.frx":2BEA
      Icon            =   "frm_addressbook.frx":3184
      Begin osenxpsuite.OsenXPButton CmdOpenBrowser 
         Height          =   315
         Left            =   4590
         TabIndex        =   49
         ToolTipText     =   "Open Browser"
         Top             =   2490
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_addressbook.frx":371E
         PICN            =   "frm_addressbook.frx":373A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         XPBlendPicture  =   0   'False
         GradientColor1  =   16772057
         GradientColor2  =   16777215
      End
      Begin osenxpsuite.OsenXPButton CmdSendMail 
         Height          =   315
         Left            =   4590
         TabIndex        =   48
         ToolTipText     =   "Send Mail"
         Top             =   2130
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_addressbook.frx":3CD4
         PICN            =   "frm_addressbook.frx":3CF0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         XPBlendPicture  =   0   'False
         GradientColor1  =   16772057
         GradientColor2  =   16777215
      End
      Begin osenxpsuite.OsenXPButton CmdRemovePicture 
         Height          =   285
         Left            =   4020
         TabIndex        =   47
         ToolTipText     =   "Remove current picture"
         Top             =   1770
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   503
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "&Remove"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_addressbook.frx":408A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         OffsetLeft      =   0
         OffsetTop       =   0
         XPBlendPicture  =   0   'False
         GradientColor1  =   16772057
         GradientColor2  =   16777215
      End
      Begin VB.PictureBox PicPhoto 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   1485
         Left            =   3360
         ScaleHeight     =   1455
         ScaleWidth      =   1545
         TabIndex        =   45
         Top             =   240
         Width           =   1575
         Begin VB.Image ImgPhoto 
            Height          =   1515
            Left            =   -30
            Stretch         =   -1  'True
            Top             =   -30
            Width           =   1605
         End
      End
      Begin osenxpsuite.OsenXPButton CmdAddPicture 
         Height          =   285
         Left            =   3360
         TabIndex        =   44
         ToolTipText     =   "&Add Picture"
         Top             =   1770
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "Add"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_addressbook.frx":40A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         OffsetLeft      =   0
         OffsetTop       =   0
         XPBlendPicture  =   0   'False
         GradientColor1  =   16772057
         GradientColor2  =   16777215
      End
      Begin osenxpsuite.OsenXPTextBox OxtData 
         Height          =   285
         Index           =   0
         Left            =   1170
         TabIndex        =   0
         Top             =   330
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         DataField       =   "Fname"
      End
      Begin osenxpsuite.OsenXPTextBox OxtData 
         Height          =   285
         Index           =   1
         Left            =   1170
         TabIndex        =   1
         Top             =   690
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         DataField       =   "MName"
      End
      Begin osenxpsuite.OsenXPTextBox OxtData 
         Height          =   285
         Index           =   2
         Left            =   1170
         TabIndex        =   2
         Top             =   1050
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         DataField       =   "Lname"
      End
      Begin osenxpsuite.OsenXPTextBox OxtData 
         Height          =   285
         Index           =   3
         Left            =   1170
         TabIndex        =   3
         Top             =   1410
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         DataField       =   "NickName"
      End
      Begin osenxpsuite.OsenXPTextBox OxtData 
         Height          =   285
         Index           =   4
         Left            =   1170
         TabIndex        =   4
         Top             =   1770
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         DataField       =   "Mobile"
      End
      Begin osenxpsuite.OsenXPTextBox OxtData 
         Height          =   285
         Index           =   5
         Left            =   1170
         TabIndex        =   5
         Top             =   2130
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   503
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         DataField       =   "email"
      End
      Begin osenxpsuite.OsenXPTextBox OxtData 
         Height          =   285
         Index           =   6
         Left            =   1170
         TabIndex        =   6
         Top             =   2490
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   503
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         DataField       =   "website"
      End
      Begin VB.Label LblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Website:"
         Height          =   195
         Index           =   6
         Left            =   150
         TabIndex        =   41
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label LblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   40
         Top             =   2160
         Width           =   420
      End
      Begin VB.Label LblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile:"
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   39
         Top             =   1800
         Width           =   510
      End
      Begin VB.Label LblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nick Name:"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   38
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label LblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name*:"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   37
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label LblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name:"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   36
         Top             =   720
         Width           =   960
      End
      Begin VB.Label LblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name*:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   35
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Scheme:"
      Height          =   195
      Left            =   240
      TabIndex        =   69
      Top             =   5760
      Width           =   1035
   End
End
Attribute VB_Name = "Frm_addressbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Rs As ADODB.Recordset
Attribute Rs.VB_VarHelpID = -1
Private MyContactId             As Long
Private IsNewRecord             As Boolean
Private MyKey                   As Integer

Private Sub CboCountry_Change(Index As Integer)

  ' auto list

    AutoListView CboCountry(Index), MyKey

End Sub

Private Sub CboCountry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    MyKey = KeyCode

End Sub

Private Sub CboScheme_Click()

  ' Change color scheme

    XPF.ColorScheme = CboScheme.ListIndex
    Drawheader

End Sub

Private Sub CmdAction_Click(Index As Integer)

    If Index = 0 Then
        If CmdAction(0).Caption = "&Add New" Then

            IsNewRecord = True
            MyData.AddNew
            SetButton False

            ' clean up picture
            Set ImgPhoto.Picture = Nothing
            DoEvents

            OxtData(0).SetFocus

          Else
            If OxtData(0).Text = "" Or OxtData(2).Text = "" Then
                MsgBoxXP "FirstName and LastName required !!!", vbExclamation, "Warning!!!", , XPF.ColorScheme
              Else
                writecontactinfo
                MyData.UpdateBatch adAffectAllChapters
                SetButton True
                If IsNewRecord Then
                    LstContact.AddItem OxtData(0).Text & " " & OxtData(2).Text
                End If
                IsNewRecord = False
            End If
        End If
      Else
        If CmdAction(0).Caption = "&Add New" Then
            MyData.Edit
            SetButton False
          Else
            If IsNewRecord Then
                MyData.Cancel
              Else
                MyData.CancelUpdate
                If Not (Rs.EOF And Rs.BOF) Then
                    ReadContactFromRs
                End If
            End If
            SetButton True
        IsNewRecord = False
        End If
    End If

End Sub

Private Sub CmdAddPicture_Click()
Dim Sfile As String
    On Error Resume Next
        Sfile = ShowOpenDialog("Open Picture files", "All Picture files|*.bmp;*.jpg;*.png;*.jpeg")
        
        ' check filename
        If Sfile <> "" Then
        On Error GoTo ErrExit
        Set ImgPhoto.Picture = LoadPicture(Sfile)
        MyData.SaveFileToRecordset Sfile, "photo", "imgsize"
    End If
ErrExit:

End Sub

Private Sub CmdDelete_Click()

    If Not (Rs.EOF And Rs.BOF) Then
        If MsgBoxXP("Are you sure to delete these record?", vbQuestion + vbYesNo, "Delete Record", , XPF.ColorScheme) = vbYes Then
            LstContact.RemoveItem Rs.AbsolutePosition - 1
            MyData.Delete
        End If
    End If

End Sub

Private Sub CmdExit_Click()

    Unload Me

End Sub

Private Sub CmdOpenBrowser_Click()

  ' Open website

    If OxtData(6).Text <> "" Then
        ShellExecute Me.hWnd, "open", OxtData(6).Text, vbNullString, vbNullString, &H1
    End If

End Sub

Private Sub CmdRefresh_Click()

    StrSQL = "Select * from tb_contact order by fname"
    Rs.Close
    OpenRecordset

End Sub

Private Sub CmdRemovePicture_Click()

    On Error Resume Next
        ' clean up picture
        Set ImgPhoto.Picture = Nothing
        ' remove picture from recordset
        Rs.Fields("photo") = Null
        Rs.Fields("imgsize") = 0

    On Error GoTo 0

End Sub

Private Sub CmdSearch_Click()

  Dim stra As String

    stra = InputBoxXP("Enter firstname or lastname which you want to search!!!", "Search....", XPF.ColorScheme)
    If stra <> "" Then
        StrSQL = "select * from tb_contact where fname like '%" & stra & "%' or lname like '%" & stra & "%' "
        Rs.Close
        OpenRecordset
    End If

End Sub

Private Sub CmdSendMail_Click()

  ' send mail

    If OxtData(5).Text <> "" Then
        ShellExecute Me.hWnd, "open", "mailto:" & OxtData(5).Text, vbNullString, vbNullString, &H3
    End If

End Sub

Private Sub Form_Load()

  ' xpform initialize

    Me.XPF.Init Me
    Drawheader

    ' init combo list
    CreateListRecord CN, "Select countryName from tb_country order by countryname", CboCountry(0)
    CreateListRecord CN, "Select countryName from tb_country order by countryname", CboCountry(1)

    ' open recordset
    Set Rs = New ADODB.Recordset
    StrSQL = "Select * from tb_contact order by fname"
    OpenRecordset
  Dim oxp As OsenXPTextBox

    For Each oxp In OxtData
        Set oxp.DataSource = Rs
        
    Next oxp
    ' setting first button
    SetButton True

End Sub

Private Sub Form_Unload(Cancel As Integer)

  ' prompt to the user to make sure exit or not ???

    If MsgBoxXP("Are you sure to exit?", vbQuestion + vbYesNo, "Exit", , XPF.ColorScheme) = vbYes Then
        ' clean up
        Rs.Close
        Set Rs = Nothing
        ' ending program <terminate>
        EndProgram
      Else
        ' cancel by user
        Cancel = 1
    End If

End Sub

Private Sub LstContact_Click()

    On Error Resume Next
        If Not (Rs.EOF And Rs.BOF) Then
            Rs.AbsolutePosition = LstContact.ListIndex + 1
        End If

End Sub

Private Sub OsenXPTab1_TabPressed(PreviousTab As Integer)

    OsenXPTab1.SetActivePage PicPage

End Sub

Private Sub Drawheader()

  ' drawgradient for header

    PicHeader.AutoRedraw = True
    PicHeader.Cls
    DrawGradients PicHeader.hDC, 0, 0, PicHeader.Width * 2 / 3, PicHeader.Height, LstContact.BackSelectedG2, LstContact.BackSelectedG1, 0
    DrawGradients PicHeader.hDC, PicHeader.Width * 2 / 3, 0, PicHeader.Width, PicHeader.Height, LstContact.BackSelectedG1, LstContact.BackSelectedG2, 0
    PicHeader.Refresh
    PicHeader.AutoRedraw = False

    ' change page backcolor
    PicPage(1).BackColor = Me.BackColor
    PicPage(2).BackColor = Me.BackColor

End Sub

Private Sub OpenRecordset()

    CreateListContact

    ' open recordset
    Rs.CursorLocation = adUseClient
    Rs.Open StrSQL, CN, adOpenStatic, adLockOptimistic

    ' binding records
    Set MyData.DataSource = Rs
    DoEvents

    ' move first
    If Not (Rs.EOF And Rs.BOF) Then
        MyData.MoveFirst
    End If

End Sub

Private Sub ReadContactFromRs()

  Dim i As Long

    On Error GoTo ErrX

'  Dim oxp As OsenXPTextBox
'
'    For Each oxp In OxtData
'        oxp.Text = ""
'    Next oxp

    CboCountry(0).Text = ""
    CboCountry(1).Text = ""

    If Not (Rs.EOF And Rs.BOF) Then
        If Not IsNull(Rs.Fields("PersonalID").Value) Then
            MyContactId = Rs.Fields("PersonalID").Value

'            ' read contact info
'            For Each oxp In OxtData
'                If Not IsNull(Rs.Fields(oxp.DataField).Value) Then
'                    oxp.Text = Rs.Fields(oxp.DataField).Value
'                End If
'            Next oxp

            If Not IsNull(Rs.Fields(CboCountry(0).DataField).Value) Then
                CboCountry(0).Text = Rs.Fields(CboCountry(0).DataField).Value
            End If

            If Not IsNull(Rs.Fields(CboCountry(1).DataField).Value) Then
                CboCountry(1).Text = Rs.Fields(CboCountry(1).DataField).Value
            End If

            ' clean up photo
            Set ImgPhoto.Picture = Nothing

            ' check & get picture from recordset
            If Not IsNull(Rs.Fields("photo").Value) Then

                ' check temp file
                StrSQL = App.Path & "\temp\test.tmp"
                If FileLen(StrSQL) > 0 Then Kill StrSQL
                DoEvents

                ' get picture from recordset and save into temp folder
                MyData.GetFileFromRecordset StrSQL, "photo", "imgsize"
                DoEvents

                ImgPhoto.Picture = LoadPicture(StrSQL)
                DoEvents

            End If

        End If

    End If
ErrX:

End Sub

Private Sub SetButton(IpNew As Boolean)

    On Error Resume Next

        If IpNew = False Then
            CmdAction(0).Caption = "&Update"
            CmdAction(1).Caption = "&Cancel"

          Else
            CmdAction(0).Caption = "&Add New"
            CmdAction(1).Caption = "&Edit"
        End If

        LstContact.Enabled = IpNew
        CmdSearch.Enabled = IpNew
        CmdDelete.Enabled = IpNew
        CmdRefresh.Enabled = IpNew
        CmdExit.Enabled = IpNew
        CmdAddPicture.Enabled = Not IpNew
        CmdRemovePicture.Enabled = Not IpNew
        MyData.Enabled = IpNew

      Dim o As OsenXPTextBox

        For Each o In OxtData
            o.Locked = IpNew
        Next o

End Sub

Private Sub CreateListContact()

    On Error GoTo Erry
  Dim R As New ADODB.Recordset
  Dim stra As String

    R.Open StrSQL, CN, adOpenStatic, adLockOptimistic
    LstContact.Clear
    If Not (R.EOF And R.BOF) Then
        Do While Not R.EOF
            stra = R.Fields("fname").Value & " " & R.Fields("lname").Value
            LstContact.AddItem stra
            R.MoveNext
        Loop
    End If

Exit Sub

Erry:
    MsgBoxXP "Error No: " & Err.Number & vbLf & _
             Err.Description, vbCritical, "Error", , XPF.ColorScheme
    Err.Clear

End Sub

Private Sub writecontactinfo()

  Dim oxp As OsenXPTextBox

    For Each oxp In OxtData
        If oxp.Text <> "" Then
            Rs.Fields(oxp.DataField).Value = oxp.Text
        End If
    Next oxp

    If CboCountry(0).Text <> "" Then
        Rs.Fields(CboCountry(0).DataField).Value = CboCountry(0).Text
    End If

    If CboCountry(1).Text <> "" Then
        Rs.Fields(CboCountry(1).DataField).Value = CboCountry(1).Text
    End If

End Sub

Private Sub OxtData_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    If Not (Rs.EOF And Rs.BOF) Then
        ReadContactFromRs
    End If

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2004-Jul-28 14:07) 5 + 397 = 402 Lines
