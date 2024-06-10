VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FormDR 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  DELIVERY RECEIPT"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16905
   ControlBox      =   0   'False
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormDR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10695
   ScaleMode       =   0  'User
   ScaleWidth      =   25196.62
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDRSearch 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1665
      Left            =   14160
      TabIndex        =   61
      Top             =   6360
      Visible         =   0   'False
      Width           =   1545
      Begin VB.TextBox txtDRSearch 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   240
         MaxLength       =   50
         TabIndex        =   62
         Top             =   270
         Width           =   7215
      End
      Begin MSComctlLib.ListView lvwDRSearch 
         Height          =   5150
         Left            =   0
         TabIndex        =   63
         Top             =   960
         Width           =   16900
         _ExtentX        =   29819
         _ExtentY        =   9075
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.TextBox txtDRTotalPcs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   8760
      Width           =   2348
   End
   Begin VB.Frame FrameInv 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   6315
      Left            =   3360
      TabIndex        =   82
      Top             =   2280
      Visible         =   0   'False
      Width           =   9435
      Begin VB.CommandButton Command1 
         Caption         =   "GO TO PRODUCT LIBRARY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   315
         TabIndex        =   83
         Top             =   7250
         Width           =   7100
      End
      Begin MSComctlLib.ListView lvwInv 
         Height          =   5805
         Left            =   240
         TabIndex        =   84
         Top             =   300
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   10239
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   12480
      TabIndex        =   69
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame FrameProduct 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1275
      Left            =   240
      TabIndex        =   53
      Top             =   5640
      Visible         =   0   'False
      Width           =   2115
      Begin VB.TextBox txtDRProductSearch 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   280
         MaxLength       =   50
         TabIndex        =   28
         Top             =   300
         Width           =   7100
      End
      Begin VB.CommandButton cmdAddProduct 
         Caption         =   "GO TO PRODUCT LIBRARY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   315
         TabIndex        =   30
         Top             =   7250
         Width           =   7100
      End
      Begin MSComctlLib.ListView lvwProduct 
         Height          =   6270
         Left            =   285
         TabIndex        =   29
         Top             =   840
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   11060
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame FrameDestination 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1905
      Left            =   14160
      TabIndex        =   66
      Top             =   2400
      Visible         =   0   'False
      Width           =   2235
      Begin VB.CommandButton cmdAddDueTo 
         Caption         =   "ADD DESTINATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   250
         TabIndex        =   77
         Top             =   7200
         Width           =   6450
      End
      Begin VB.CommandButton cmdAddVehicle 
         Caption         =   "ADD VEHICLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   250
         TabIndex        =   70
         Top             =   6600
         Width           =   6450
      End
      Begin VB.TextBox txtDRDriver 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         MaxLength       =   50
         TabIndex        =   17
         Top             =   360
         Width           =   6450
      End
      Begin VB.TextBox txtVehicleSearch 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         MaxLength       =   50
         TabIndex        =   18
         Top             =   800
         Width           =   6450
      End
      Begin VB.TextBox txtDestinationSearch 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   250
         MaxLength       =   50
         TabIndex        =   20
         Top             =   3720
         Width           =   6450
      End
      Begin MSComctlLib.ListView lvwDestination 
         Height          =   2235
         Left            =   240
         TabIndex        =   21
         Top             =   4200
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   3942
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwVehicle 
         Height          =   2310
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   4075
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame frameDelivered 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1695
      Left            =   14160
      TabIndex        =   45
      Top             =   4440
      Visible         =   0   'False
      Width           =   2235
      Begin VB.ComboBox txtDRSaleTo 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "FormDR.frx":08CA
         Left            =   240
         List            =   "FormDR.frx":08DD
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   320
         Width           =   6450
      End
      Begin VB.CommandButton cmdAddDeliver 
         Caption         =   "ADD CUSTOMER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   250
         TabIndex        =   22
         Top             =   7320
         Width           =   6450
      End
      Begin VB.TextBox txtDeliverSearch 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   250
         MaxLength       =   50
         TabIndex        =   15
         Top             =   750
         Width           =   6450
      End
      Begin MSComctlLib.ListView lvwDelivered 
         Height          =   5955
         Left            =   255
         TabIndex        =   16
         Top             =   1200
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   10504
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame FrameDRDetails 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   2505
      Left            =   240
      TabIndex        =   46
      Top             =   2400
      Visible         =   0   'False
      Width           =   2715
      Begin VB.ComboBox txtDRPhase 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         ItemData        =   "FormDR.frx":092B
         Left            =   2520
         List            =   "FormDR.frx":0944
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   2400
         Width           =   4100
      End
      Begin MSComctlLib.ListView lvwPhase 
         Height          =   495
         Left            =   1080
         TabIndex        =   80
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwBlock 
         Height          =   495
         Left            =   1080
         TabIndex        =   79
         Top             =   960
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   873
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.TextBox txtDRBlock 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         MaxLength       =   50
         TabIndex        =   78
         Top             =   1440
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtDRUnit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   75
         Top             =   5520
         Width           =   1455
      End
      Begin VB.TextBox txtDRArea 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtDRQty 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         MaxLength       =   13
         TabIndex        =   31
         Top             =   4080
         Width           =   4100
      End
      Begin VB.CommandButton cmdSaveDRDetails 
         Caption         =   "S A V E"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   7080
         Width           =   4100
      End
      Begin VB.TextBox txtDRAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   2500
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   6240
         Width           =   4100
      End
      Begin VB.TextBox txtDRCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2500
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   33
         Top             =   5520
         Width           =   2535
      End
      Begin VB.TextBox txtDRWeight 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         MaxLength       =   13
         TabIndex        =   32
         Top             =   4800
         Width           =   4100
      End
      Begin VB.TextBox txtDRProduct 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2520
         TabIndex        =   25
         Top             =   1750
         Width           =   4100
      End
      Begin VB.TextBox txtDRSackBox 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         MaxLength       =   13
         TabIndex        =   27
         Top             =   3360
         Width           =   4100
      End
      Begin VB.CheckBox checkCOntracted 
         BackColor       =   &H80000000&
         Caption         =   "CONTRACTED"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         TabIndex        =   24
         Top             =   1200
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "AREA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   480
         TabIndex        =   68
         Top             =   2640
         Width           =   675
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Left            =   2500
         TabIndex        =   64
         Top             =   1080
         Width           =   4100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "QUANTITY"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   495
         TabIndex        =   55
         Top             =   4250
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "DELIVERY RECEIPT DETAILS"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   13
         Left            =   1680
         TabIndex        =   52
         Top             =   345
         Width           =   3600
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   360
         Picture         =   "FormDR.frx":098E
         Stretch         =   -1  'True
         Top             =   300
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "AMOUNT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   495
         TabIndex        =   51
         Top             =   6360
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "UNIT PRICE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   495
         TabIndex        =   50
         Top             =   5700
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "WEIGHT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   480
         TabIndex        =   49
         Top             =   5000
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "PRODUCT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   500
         TabIndex        =   48
         Top             =   1920
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "SACK/BOX"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   495
         TabIndex        =   47
         Top             =   3480
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   240
         TabIndex        =   54
         Top             =   240
         Width           =   6400
      End
   End
   Begin VB.TextBox txtDRTotalCuM 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   14400
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   8750
      Width           =   2348
   End
   Begin VB.TextBox txtDRTotalBdFt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   12000
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   8750
      Width           =   2348
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000008&
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
      Height          =   850
      Left            =   0
      TabIndex        =   42
      Top             =   9600
      Width           =   16874
      Begin VB.CommandButton cmdReport 
         Caption         =   "REPORT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   10150
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   2000
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&XIT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   14650
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   2000
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "CANCEL"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   6130
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   2000
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "PRINT "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   2000
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "ADD"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   2000
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4130
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   2000
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&NEW"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   2000
      End
   End
   Begin MSComctlLib.ListView lvwDR 
      Height          =   6300
      Left            =   0
      TabIndex        =   41
      Top             =   2300
      Width           =   16905
      _ExtentX        =   29819
      _ExtentY        =   11113
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1600
      Left            =   0
      TabIndex        =   36
      Top             =   700
      Width           =   16907
      Begin VB.TextBox txtCHINum 
         BackColor       =   &H80000003&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5600
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1050
         Width           =   2500
      End
      Begin VB.TextBox txtARNum 
         BackColor       =   &H80000003&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5600
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   200
         Width           =   2500
      End
      Begin VB.TextBox txtCINum 
         BackColor       =   &H80000003&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5600
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   2500
      End
      Begin VB.TextBox txtSONum 
         BackColor       =   &H80000003&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1450
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1050
         Width           =   2500
      End
      Begin VB.TextBox txtDRDestination 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10200
         TabIndex        =   14
         Top             =   600
         Width           =   6300
      End
      Begin VB.TextBox txtDRNum 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1450
         MaxLength       =   15
         TabIndex        =   7
         Top             =   200
         Width           =   2500
      End
      Begin VB.TextBox txtDRDelivered 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10200
         TabIndex        =   13
         Top             =   200
         Width           =   6300
      End
      Begin VB.TextBox txtDRRemarks 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10200
         MaxLength       =   60
         TabIndex        =   23
         Top             =   1050
         Width           =   6300
      End
      Begin MSMask.MaskEdBox txtDRDate 
         Height          =   420
         Left            =   1450
         TabIndex        =   8
         Top             =   600
         Width           =   2500
         _ExtentX        =   4392
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblCHI 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "CH.I. #"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4600
         TabIndex        =   74
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblCI 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "C.I. #"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4600
         TabIndex        =   73
         Top             =   675
         Width           =   480
      End
      Begin VB.Label lblAR 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "A.R. #"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4600
         TabIndex        =   72
         Top             =   255
         Width           =   555
      End
      Begin VB.Label lblSO 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "REF. #"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   255
         TabIndex        =   71
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "D.R. DATE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   255
         TabIndex        =   67
         Top             =   675
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "TRANSPO"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   14
         Left            =   8800
         TabIndex        =   65
         Top             =   675
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "D.R. NUM"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   255
         TabIndex        =   39
         Top             =   255
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "RECEIVER"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   8800
         TabIndex        =   38
         Top             =   250
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "REMARKS"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   8800
         TabIndex        =   37
         Top             =   1100
         Width           =   945
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "PIECES"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   11280
      TabIndex        =   86
      Top             =   9240
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "CU.MT."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   21
      Left            =   15600
      TabIndex        =   60
      Top             =   9255
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "TOTALS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   8280
      TabIndex        =   58
      Top             =   8880
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "BOARD FT."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   17
      Left            =   13320
      TabIndex        =   57
      Top             =   9240
      Width           =   1035
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Header"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   840
      TabIndex        =   44
      Top             =   360
      Width           =   690
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPANY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   840
      TabIndex        =   43
      Top             =   120
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   0
      Picture         =   "FormDR.frx":4015
      Stretch         =   -1  'True
      Top             =   45
      Width           =   738
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   -120
      TabIndex        =   40
      Top             =   0
      Width           =   17355
   End
End
Attribute VB_Name = "FormDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn             As ADODB.Connection
Private mmsAdoCmd              As ADODB.Command
Private mmsADORst              As ADODB.Recordset
Public strsql, EncodeMode       As String

Private DRLI       As ListItem
Private ItemEdit   As Integer

Private InvID, DRID, DRBuyer, DRTranspo, DRRemarks          As String
Private StartMonth, EndMonth                                As String
Private DRTotalQty, DRItemNo, DRTotalBdFt, DRTotalCuM, DRTotalPcs       As Long
Private Sub Command2_Click()
        'strsql = "Update DRDetails SET DRARNum = '" & "-" & "'"
        'strsql = strsql & ", DRSONum = '" & "-" & "'"
        'strsql = strsql & ", DRCINum = '" & "-" & "'"
        'strsql = strsql & ", DRCHINum = '" & "-" & "'"
        'strsql = strsql & ", DRNumDetails = '" & "S.O.#- / A.R.#- / C.I.#- / CH.I.#-" & "'"
        'strsql = strsql & ", DRPhase = '" & "PH-" & "'"
        'strsql = strsql & ", DRBlock = '" & "BLK-" & "'"
        'strsql = strsql & ", DRAreaDetails = '" & "PH- / BLK-" & "'"
        'CommandExecute
        
        'StartMonth = "01/01/2019"
        'EndMonth = "12/31/2019"
        
        'strsql = "UPDATE DRDetails  SET DRCINum = DRRemarks " _
        '      & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#"
        'CommandExecute
        
        'strsql = "Update DRDetails SET DRClass = '" & "MIX GRADE" & "' WHERE DRClass like '" & "MIX GRADES" & "'"
        'CommandExecute

        'strsql = "Update DRDetails SET DRCustomer = '" & "JO MAGUINSAY" & "' WHERE DRClass like '" & "SPOILAGE" & "' AND DRRemarks like '" & "MAGUINSAY SPOILAGE" & "' "
        'CommandExecute
End Sub
Private Sub Form_Load()
   EncodeMode = "START"
   ConnectToDB
   SetlvwDR
   lblComp.Caption = FormMainMenu.lblComp.Caption
   lblHeader.Caption = FormMainMenu.lblHeader.Caption
End Sub
Private Sub GetFrom()
   mmsAdoCmd.CommandText = "Select * from Settings"
   Set mmsADORst = mmsAdoCmd.Execute
   'txtDRFrom.Text = mmsADORst.Fields("Area") & "-" & mmsADORst.Fields("AreaLoc")
End Sub
Private Sub txtDRDate_GotFocus()
   txtDRDate.SelLength = Len(txtDRDate.Text)
End Sub
Private Sub TxtDRdate_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then
      Exit Sub
   ElseIf Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 Then
      SendKeys_
   ElseIf KeyAscii = 13 Then
     If Not IsDate(txtDRDate.Text) Then
        MsgBox "Invalid Date.", vbExclamation, "Invalid Entry"
        txtDRDate.SetFocus
        txtDRDate = Format$(Now, "mm/dd/yyyy")
     Else
        txtDRDelivered.SetFocus
     End If
   End If
End Sub
Private Sub txtDRDelivered_DblClick()
    SetDelivered
    BoxState False
    LoadDelivered
    txtDRSaleTo.Text = "-"
End Sub
Private Sub txtDRDelivered_GotFocus()
   txtDRDelivered.SelLength = Len(txtDRDelivered.Text)
End Sub
Private Sub txtDRDelivered_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = vbKeyEscape Then
      Exit Sub
   ElseIf KeyAscii > 0 Then
      txtDRDelivered_DblClick
   End If
End Sub
Private Sub txtDRSaleTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtDeliverSearch.SetFocus
End If
End Sub
Private Sub txtDRDestination_DblClick()
    SetDestination
    BoxState False
    LoadVehicle
    LoadDestination
    txtDRDriver.SetFocus
End Sub
Private Sub txtDRDestination_GotFocus()
   txtDRDestination.SelLength = Len(txtDRDestination.Text)
End Sub
Private Sub txtDRDestination_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = vbKeyEscape Then
      Exit Sub
   ElseIf KeyAscii > 0 Then
      txtDRDestination_DblClick
   End If
End Sub
Private Sub txtDRDriver_GotFocus()
  txtDRDriver.SelLength = Len(txtDRDriver.Text)
End Sub
Private Sub txtDRDriver_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = vbKeyEscape Then
      Exit Sub
   ElseIf KeyAscii = 13 Then
      txtVehicleSearch.SetFocus
   End If
End Sub
Private Sub txtDRRemarks_GotFocus()
  txtDRRemarks.SelLength = Len(txtDRRemarks.Text)
End Sub
Private Sub txtDRRemarks_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = vbKeyEscape Then
      Exit Sub
   ElseIf KeyAscii = 13 Then
      DRBuyer = txtDRDelivered.Text: DRTranspo = txtDRDestination.Text: DRRemarks = txtDRRemarks.Text
      LoadInventory
   End If
End Sub
Private Sub LoadInventory()
SetInventory
strsql = "SELECT * from InventoryTemp ORDER By InvID": CommandExecute
    lvwInv.ListItems.Clear
    With mmsADORst
        Do Until .EOF
        Set DRLI = lvwInv.ListItems.Add(, , !HDate & "")
            DRLI.SubItems(1) = !Area & ""
            DRLI.SubItems(2) = !Specie & ""
            DRLI.SubItems(3) = !Hill & ""
            DRLI.SubItems(4) = !Bolt & ""
            DRLI.SubItems(5) = !Size & ""
            DRLI.SubItems(6) = !Pcs & ""
            DRLI.SubItems(7) = !BdFt & ""
            DRLI.SubItems(8) = !InvID & ""
            .MoveNext
        Loop
     End With
End Sub
'----------------------------------------------------------------------------

'----------------------------------------------------------------------------
Private Sub lvwInv_DblClick()
    FrameInv.Visible = False
    SaveDRDetailsTemp
        'delete on inv
        InvID = lvwInv.SelectedItem.SubItems(8)
        strsql = "DELETE FROM InventoryTemp WHERE InvID = " & InvID & "": CommandExecute
        'compute boardfeet/cum
        strsql = " SELECT SUM(DRBdFt) as SubBdFt, SUM(DRPcs) as SubPcs  FROM DRDetailsTemp ": CommandExecute
        DRTotalPcs = mmsADORst.Fields!SubPcs: txtDRTotalPcs.Text = DRTotalPcs
        DRTotalBdFt = mmsADORst.Fields!SubBdFt: txtDRTotalBdFt.Text = DRTotalBdFt
        txtDRTotalCuM.Text = CDbl(CDbl(DRTotalBdFt) / 423.776)
   
   InsertDRTotals
   cmdPrint.Enabled = True
End Sub

'----------------  D E L I V E R   S E A R C H ----------------------
Private Sub txtDeliverSearch_Change()
    strsql = "Select * from Delivered where Customer like '" & txtDeliverSearch.Text & "%' Order by Customer": CommandExecute
    lvwDelivered.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set DRLI = lvwDelivered.ListItems.Add(, , !Customer & "")
            .MoveNext
        Loop
     End With
End Sub
Private Sub txtDeliverSearch_Click()
    txtDRDelivered.Text = ""
End Sub
Private Sub txtDeliverSearch_GotFocus()
   txtDeliverSearch.SelLength = Len(txtDeliverSearch.Text)
End Sub
Private Sub txtDeliverSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      lvwDelivered.SetFocus
   End If
End Sub
'----------------  V E H I C L E   S E A R C H ----------------------
Private Sub txtVehicleSearch_Change()
strsql = "Select * from Vehicle where VehicleName like '" & txtVehicleSearch.Text & "%' Order by VehicleName": CommandExecute
    lvwVehicle.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set DRLI = lvwVehicle.ListItems.Add(, , !VehicleName & "")
                DRLI.SubItems(1) = !PlateNo & ""
                .MoveNext
        Loop
     End With
End Sub
Private Sub txtVehicleSearch_Click()
    txtVehicleSearch.Text = ""
End Sub
Private Sub txtVehicleSearch_GotFocus()
  txtVehicleSearch.SelLength = Len(txtVehicleSearch.Text)
End Sub
Private Sub txtVehicleSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      lvwVehicle.SetFocus
   End If
End Sub
'----------------  D U E  T O   S E A R C H ----------------------
Private Sub txtDestinationSearch_Change()
strsql = "Select * from Destination where Destination like '" & txtDestinationSearch.Text & "%' Order by Destination": CommandExecute
    lvwDestination.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set DRLI = lvwDestination.ListItems.Add(, , !Destination & "")
            .MoveNext
        Loop
     End With
End Sub
Private Sub txtDestinationSearch_Click()
    txtDRDestination.Text = ""
End Sub
Private Sub txtDestinationSearch_GotFocus()
  txtDestinationSearch.SelLength = Len(txtDestinationSearch.Text)
End Sub
Private Sub txtDestinationSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      lvwDestination.SetFocus
   End If
End Sub
'----------------  P R O D U C T   S E A R C H ----------------------
Private Sub txtDRProductSearch_Change()
    strsql = "Select * from Inventory where Specie like '" & txtDRProductSearch.Text & "%' Order by InvID": CommandExecute
    lvwProduct.ListItems.Clear
    With mmsADORst
        Do Until .EOF
        Set DRLI = lvwProduct.ListItems.Add(, , !Date & "")
            DRLI.SubItems(1) = !Specie & ""
            DRLI.SubItems(2) = !Hill & ""
            DRLI.SubItems(3) = !Bolt & ""
            DRLI.SubItems(4) = !Size & ""
            DRLI.SubItems(5) = !Pcs & ""
            DRLI.SubItems(6) = !BdFt & ""
            .MoveNext
        Loop
     End With
End Sub
Private Sub txtDRProductSearch_Click()
    txtDRProduct.Text = ""
End Sub
Private Sub txtDRProductSearch_GotFocus()
   txtDRProductSearch.SelLength = Len(txtDRProductSearch.Text)
End Sub
Private Sub txtDRProductSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      lvwProduct.SetFocus
   End If
End Sub
'--------------------------------------------------------------------------
'                               B U T T O N S
'--------------------------------------------------------------------------
Private Sub CmdNew_Click()
Dim DRGroup As String
    EncodeMode = "A": BoxState True: ButtonState False
    ClearBox
    ClearFrame
     cmdAdd.Enabled = True: cmdCancel.Enabled = True: lvwDR.ListItems.Clear
     DRGroup = FormMainMenu.lblArea.Caption & "-" & "DR"
     DRID = GetNextDRID
     txtDRNum.Text = DRGroup & "-" & Format$(GetNextDRID, "000000")
     DRItemNo = 0
     txtDRDate = Format$(Now, "mm/dd/yyyy")
     txtDRDate.SetFocus
     GetFrom
     DeleteTemporary
End Sub
Private Sub CmdAdd_Click()
   If Not DataValidation Then
       Exit Sub
   End If
   EncodeMode = "A"
   LoadInventory
      If lvwInv.ListItems.Count = 0 Then
        MsgBox "No stocks on inventory"
        FrameInv.Visible = False
        cmdPrint.SetFocus
      End If
End Sub
Private Sub DeleteTemporary()
On Error GoTo LocalError
    strsql = "Delete From DRDetailsTemp": CommandExecute
    strsql = "DELETE * FROM InventoryTemp": CommandExecute
        strsql = "INSERT INTO InventoryTemp SELECT * FROM Inventory": CommandExecute
LocalError:
    Exit Sub
End Sub
'--------------------------------------------------------------------------------------
'                          S  E  A  R  C  H
'--------------------------------------------------------------------------------------
Private Sub cmdSearch_Click()
   EncodeMode = "S"
   DeleteTemporary
   ClearBox
   ClearFrame
   BoxState False
   ButtonState False
    cmdNew.Enabled = True
    cmdCancel.Enabled = True
   SetSearch
    
  strsql = "SELECT Distinct DRNum, DRDate, DRBuyer, DRTotalBdFt, DRCum, DRRemarks from DRDetails ORDER BY DRNum": CommandExecute
  LoadDRSearch
  txtDRSearch.SetFocus
    'DeleteTemporary
    
End Sub
Private Sub LoadDRSearch()
On Error GoTo LocalError
    lvwDRSearch.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set DRLI = lvwDRSearch.ListItems.Add(, , !DRNum & "")
                DRLI.SubItems(1) = !DRDate & ""
                DRLI.SubItems(2) = !DRBuyer & ""
                DRLI.SubItems(3) = !DRTotalBdFt & ""
                DRLI.SubItems(4) = !DRCum & ""
                DRLI.SubItems(5) = !DRRemarks & ""
                .MoveNext
        Loop
     End With
   
LocalError:
    Exit Sub
End Sub
'----------------  D R   S E A R C H ---------------------------
'-----------------------------------------------------------------
Private Sub txtDRSearch_Change()
    strsql = "SELECT Distinct DRNum, DRDate,DRBuyer, DRTotalBdFt, DRCum, DRRemarks from DRDetails where DRID like  '" & txtDRSearch.Text & "%'  ORDER BY DRNum"
    CommandExecute
    LoadDRSearch
End Sub
Private Sub txtDRSearch_Click()
    txtDRSearch.Text = ""
End Sub
Private Sub txtDRSearch_GotFocus()
   txtDRSearch.Text = ""
   txtDRSearch.SelLength = Len(txtDeliverSearch.Text)
End Sub
Private Sub txtDRSearch_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
      'cmdAddDeliver_Click
   End If
End Sub
Private Sub txtDRSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      lvwDRSearch.SetFocus
   End If
End Sub
'--------------------------------------------------------------------------------------
'                                   C  A  N  C  E  L
'--------------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
On Error GoTo LocalError
  
  If EncodeMode = "A" Or EncodeMode = "S" Then   '  ------- A D D / S E A R C H
    If MsgBox("Are you sure you want to CANCEL?", vbYesNo + vbQuestion, "Exit") = vbYes Then
        ClearFrame
        ClearBox
        'ClearDRDetails
        BoxState False
        ButtonState False
        cmdSearch.Enabled = True: cmdNew.Enabled = True: cmdNew.SetFocus
    Else
        Exit Sub
    End If
  End If
  
  If EncodeMode = "SF" Then   '  ------- S E A R C H  F O U N D
     If MsgBox("You want to CANCEL this DR_Number?", vbYesNo + vbQuestion, "Exit") = vbYes Then
        FormDR.Enabled = False
        FormLog.txtEncodeMode.Text = "SF"
        FormLog.Show
     Else
        Exit Sub
     End If
  End If

LocalError:
    Exit Sub
End Sub
Public Sub CancelDR()
   EncodeMode = "CANCEL"

        strsql = "Update DRDetails SET DRRemarks = '" & "CANCEL" & "'"
        strsql = strsql & ", DRQty = '" & "0.00" & "'"
        strsql = strsql & ", DRWeight = '" & "0.00" & "'"
        strsql = strsql & ", DRAmount = '" & "0.00" & "'"
        strsql = strsql & ", DRCoShare = '" & "0.00" & "'"
        strsql = strsql & ", DRTotalQty = '" & "0.00" & "'"
        strsql = strsql & ", DRTotalWeight = '" & "0.00" & "'"
        strsql = strsql & ", DRTotalAmount = '" & "0.00" & "'"
        strsql = strsql & ", DRTotalCoShare = '" & "0.00" & "'"
        strsql = strsql & " where DRNum like '" & txtDRNum.Text & "'"
         CommandExecute
         Set mmsADORst = Nothing
   ClearFrame
   ClearBox
   ClearDRDetails
   BoxState False
   ButtonState False
   cmdNew.Enabled = True
   cmdSearch.Enabled = True
   cmdNew.SetFocus

End Sub
Private Sub cmdPrint_Click()
   If EncodeMode = "A" Or EncodeMode = "E" Then
      strsql = " Insert Into DRDetails Select * From DRDetailsTemp "
      CommandExecute
   End If

   ClearFrame
   ClearBox
   ClearDRDetails
   BoxState False
   ButtonState False
   cmdNew.Enabled = True
   cmdSearch.Enabled = True
   Load DataEnvironment1
   If DataEnvironment1.rsCommand8.State <> 0 Then DataEnvironment1.rsCommand8.Close
     ReportDR.Refresh
   If ReportDR.Visible = False Then ReportDR.Show
   
   If EncodeMode = "SF" Then
      DeleteTemporary
      'ReportDR.Sections("Pageheader").Controls("Label13").Caption = "DR"
   End If
Set mmsADORst = Nothing
End Sub
Private Sub cmdReport_Click()
   FormDRReport.Show
   Form_Load
   Unload Me
End Sub
Private Sub cmdExit_Click()
   If MsgBox("Are you sure you want to exit?", _
        vbYesNo + vbQuestion, "Exit") = vbYes Then
     ClearBox
     ClearFrame
     Unload Me
     FormMainMenu.Show
     FormMainMenu.cmdDR.SetFocus
   Else
     Exit Sub
 End If
End Sub
Private Sub cmdAddDeliver_Click()
       FormDelivered.Show
End Sub
Private Sub cmdAddProduct_Click()
       FormDR.Enabled = False
       FormLog.txtEncodeMode.Text = "PRODUCT2"
       FormLog.Show
End Sub
Private Sub cmdAddDueTo_Click()
       FormDestination.Show
End Sub
Private Sub cmdAddVehicle_Click()
       FormTranspo.Show
End Sub
'--------------------------------------------------------------------------
'                     L I S T V I E W    E V E N T S
'-------------------------------------------------------------------------

Private Sub lvwDR_DblClick()
  If lvwDR.Enabled = True Then
   EncodeMode = "E"
   ItemEdit = lvwDR.SelectedItem.Text
   LoadDRItemEdit
  End If
End Sub
'---------- D R   S E A R C H -----
Private Sub lvwDRSearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
        LoadHeader
    End With
End Sub
Private Sub lvwDRSearch_DblClick()
    FrameDRSearch.Visible = False
    BoxState False
    SaveDRDetailsTemp
    LoadDRDetails
    cmdPrint.Enabled = True
End Sub
Private Sub lvwDRSearch_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
    'cmdAddDeliver_Click
   End If
End Sub
Private Sub lvwDRSearch_KeyPress(KeyAscii As Integer)
    lvwDRSearch_DblClick
End Sub
Private Sub lvwDRSearch_GotFocus()
On Error GoTo LocalError
        LoadHeader
LocalError:
    Exit Sub
End Sub
Private Sub LoadHeader()
      txtDRNum.Text = lvwDRSearch.SelectedItem.Text
      txtDRDate.Text = Format$(lvwDRSearch.SelectedItem.SubItems(1), "mm/dd/yyyy")
      txtDRDelivered.Text = lvwDRSearch.SelectedItem.SubItems(2)
      txtDRRemarks.Text = lvwDRSearch.SelectedItem.SubItems(4)
End Sub

'---------- D E L I V E R Y -----
Private Sub lvwDelivered_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
      txtDRDelivered.Text = txtDRSaleTo.Text & " - " & lvwDelivered.SelectedItem.Text
    End With
    DRBuyer = lvwDelivered.SelectedItem.Text
End Sub
Private Sub lvwDelivered_DblClick()
    frameDelivered.Visible = False
    BoxState True
    txtDRDestination.SetFocus
End Sub
Private Sub lvwDelivered_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
    cmdAddDeliver_Click
   End If
End Sub
Private Sub lvwDelivered_KeyPress(KeyAscii As Integer)
    lvwDelivered_DblClick
End Sub
Private Sub lvwDelivered_GotFocus()
On Error GoTo LocalError
      txtDRDelivered.Text = lvwDelivered.SelectedItem.Text
      DRBuyer = lvwDelivered.SelectedItem.Text
LocalError:
    Exit Sub
End Sub
'----------   V E H I C L E   -----
Private Sub lvwVehicle_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
      txtVehicleSearch.Text = lvwVehicle.SelectedItem.Text & " (" & lvwVehicle.SelectedItem.SubItems(1) & ")"
    End With
End Sub
Private Sub lvwVehicle_DblClick()
    'FrameDestination.Visible = False
    'BoxState True
     txtDestinationSearch.SetFocus
End Sub
Private Sub lvwVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
    cmdAddDueTo_Click
   End If
End Sub
Private Sub lvwVehicle_KeyPress(KeyAscii As Integer)
    
    lvwVehicle_DblClick
End Sub
Private Sub lvwVehicle_GotFocus()
On Error GoTo LocalError
      txtVehicleSearch.Text = lvwVehicle.SelectedItem.Text & " " & lvwVehicle.SelectedItem.SubItems(1)
LocalError:
    Exit Sub
End Sub
'----------   D U E   TO   -----
Private Sub lvwDestination_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
      txtDestinationSearch.Text = lvwDestination.SelectedItem.Text
    End With
End Sub
Private Sub lvwDestination_DblClick()
    FrameDestination.Visible = False
    BoxState True
    txtDRDestination.Text = txtDRDriver.Text & " " & txtVehicleSearch.Text & "-" & txtDestinationSearch.Text
    txtDRRemarks.SetFocus
    txtDRRemarks.Text = "-"
End Sub
Private Sub lvwDestination_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
    cmdAddDueTo_Click
   End If
End Sub
Private Sub lvwDestination_KeyPress(KeyAscii As Integer)
    lvwDestination_DblClick
End Sub
Private Sub lvwDestination_GotFocus()
On Error GoTo LocalError
      txtDestinationSearch.Text = lvwDestination.SelectedItem.Text
LocalError:
    Exit Sub
End Sub
'----------   P H A S E  -----
Private Sub lvwPhase_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
      txtDRPhase.Text = lvwPhase.SelectedItem.Text
    End With
End Sub
Private Sub lvwPhase_DblClick()
    lvwPhase.Visible = False
    'BoxState True
    txtDRPhase.Text = lvwPhase.SelectedItem.Text
    txtDRSackBox.SetFocus
End Sub
Private Sub lvwPhase_KeyPress(KeyAscii As Integer)
    lvwPhase_DblClick
End Sub
Private Sub lvwPhase_GotFocus()
On Error GoTo LocalError
      txtDRPhase.Text = lvwPhase.SelectedItem.Text
LocalError:
    Exit Sub
End Sub
'----------   B L O C K  -----
Private Sub lvwBlock_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
      txtDRBlock.Text = lvwBlock.SelectedItem.Text
    End With
End Sub
Private Sub lvwBlock_DblClick()
    lvwBlock.Visible = False
    'BoxState True
    txtDRBlock.Text = lvwBlock.SelectedItem.Text
    txtDRSackBox.SetFocus
End Sub
Private Sub lvwBlock_KeyPress(KeyAscii As Integer)
    lvwBlock_DblClick
End Sub
Private Sub lvwBlock_GotFocus()
On Error GoTo LocalError
      txtDRBlock.Text = lvwBlock.SelectedItem.Text
LocalError:
    Exit Sub
End Sub
'---------- P R O D U C T -----
Private Sub lvwProduct_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
      txtDRProduct.Text = lvwProduct.SelectedItem.Text & " - " & lvwProduct.SelectedItem.SubItems(1)
      txtDRCost.Text = lvwProduct.SelectedItem.SubItems(2)
      txtDRUnit.Text = lvwProduct.SelectedItem.SubItems(3)
    End With
    DRProduct = lvwProduct.SelectedItem.Text
    DRClass = lvwProduct.SelectedItem.SubItems(1)
End Sub
Private Sub lvwProduct_DblClick()
    FrameProduct.Visible = False
    txtDRPhase.SetFocus
End Sub
Private Sub lvwProduct_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
    cmdAddDeliver_Click
   End If
End Sub
Private Sub lvwProduct_KeyPress(KeyAscii As Integer)
    lvwProduct_DblClick
End Sub
Private Sub lvwProduct_GotFocus()
On Error GoTo LocalError
      txtDRProduct.Text = lvwProduct.SelectedItem.Text & " - " & lvwProduct.SelectedItem.SubItems(1)
      txtDRCost.Text = lvwProduct.SelectedItem.SubItems(2)
      txtDRUnit.Text = lvwProduct.SelectedItem.SubItems(3)
      DRProduct = lvwProduct.SelectedItem.Text
      DRClass = lvwProduct.SelectedItem.SubItems(1)
LocalError:
    Exit Sub
End Sub
'--------------------------------------------------------------------------
'                                 F U N C T I O N S
'--------------------------------------------------------------------------
Private Sub ConnectToDB()
    Set mmsADOConn = New ADODB.Connection
    mmsADOConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MMS2.mdb" & ";Persist Security Info=False"
    mmsADOConn.Open
    
        strsql = "select * from DRDetailsTemp"
        Set mmsADORst = New ADODB.Recordset
        mmsADORst.CursorLocation = adUseClient
        mmsADORst.Open strsql, mmsADOConn, adOpenDynamic, adLockPessimistic, adCmdText
      

    Set mmsAdoCmd = New ADODB.Command
    Set mmsAdoCmd.ActiveConnection = mmsADOConn
    mmsAdoCmd.CommandType = adCmdText

End Sub
Private Sub CommandExecute()
         mmsAdoCmd.CommandText = strsql
         Set mmsADORst = mmsAdoCmd.Execute
End Sub
Private Sub LoadDelivered()
On Error GoTo LocalError
strsql = "SELECT * from Delivered ORDER BY Customer": CommandExecute
    lvwDelivered.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set DRLI = lvwDelivered.ListItems.Add(, , !Customer & "")
                .MoveNext
        Loop
     End With
LocalError:
    Exit Sub
End Sub
Private Sub LoadVehicle()
On Error GoTo LocalError
strsql = "SELECT * from Vehicle ORDER BY VehicleName": CommandExecute
    lvwVehicle.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set DRLI = lvwVehicle.ListItems.Add(, , !VehicleName & "")
                DRLI.SubItems(1) = !PlateNo & ""
                .MoveNext
        Loop
     End With
LocalError:
    Exit Sub
End Sub
Private Sub LoadDestination()
On Error GoTo LocalError
strsql = "SELECT * from Destination ORDER BY Destination": CommandExecute
    lvwDestination.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set DRLI = lvwDestination.ListItems.Add(, , !Destination & "")
                .MoveNext
        Loop
     End With
LocalError:
    Exit Sub
End Sub
Private Sub SaveDRDetailsTemp()
On Error GoTo LocalError
    If EncodeMode = "A" Then
         DRItemNo = DRItemNo + 1
         strsql = "INSERT INTO DRDetailsTemp (DRId, DRRef, DRNum, DRDate,DRBuyer, DRTranspo, DRRemarks "
         strsql = strsql & ", DRItemNo, DRHDate, DRArea, DRSpecie, DRHill, DRBolt, DRSize, DRPcs, DRBdFt) "
         strsql = strsql & "  VALUES ( " & DRID & ", " & DRID & " "
         strsql = strsql & ", '" & txtDRNum.Text & "' "
         strsql = strsql & ", '" & Replace$(txtDRDate.Text, "'", "''") & "' "
         strsql = strsql & ", '" & DRBuyer & "', '" & DRTranspo & "', '" & DRRemarks & "', " & DRItemNo & " "
         strsql = strsql & ", '" & lvwInv.SelectedItem & "' "
         strsql = strsql & ", '" & lvwInv.SelectedItem.SubItems(1) & "' "
         strsql = strsql & ", '" & lvwInv.SelectedItem.SubItems(2) & "' "
         strsql = strsql & ", '" & lvwInv.SelectedItem.SubItems(3) & "' "
         strsql = strsql & ", '" & lvwInv.SelectedItem.SubItems(4) & "' "
         strsql = strsql & ", '" & lvwInv.SelectedItem.SubItems(5) & "' "
         strsql = strsql & ", '" & lvwInv.SelectedItem.SubItems(6) & "' "
         strsql = strsql & ", '" & lvwInv.SelectedItem.SubItems(7) & "' "
         strsql = strsql & ")"
    CommandExecute
    LoadDRDetails
 End If
 
 'If EncodeMode = "E" Then
 '        strsql = "Update DRDetailsTemp SET DRContract = '" & Contracted & "'"
 '        strsql = strsql & ", DRSackBox = '" & txtDRSackBox.Text & "'"
 '        strsql = strsql & ", DRProduct = '" & DRProduct & "'"
 '        strsql = strsql & ", DRClass = '" & DRClass & "'"
 '        strsql = strsql & ", DRProductClass = '" & txtDRProduct.Text & "'"
 '        strsql = strsql & ", DRQty = '" & Format$(txtDRQty.Text, "#,###.#0") & "'"
 '        strsql = strsql & ", DRWeight= '" & txtDRWeight.Text & "'"
 '        strsql = strsql & ", DRCost = '" & Format$(txtDRCost.Text, "#,###.#0") & "'"
 '        strsql = strsql & ", DRAmount = '" & Format$(txtDRAmount.Text, "#,###.#0") & "'"
 '        strsql = strsql & ", DRCoShare = '" & Format$(DRCoShare, "#,###.#0") & "'"
 '        strsql = strsql & " Where DRItemNo like '" & ItemEdit & "'"
 '        CommandExecute
 '        Set mmsADORst = Nothing
 '  EncodeMode = "A"
 'End If
 
  If EncodeMode = "S" Then
     strsql = " Insert Into DRDetailsTemp Select * From DRDetails Where DRNum like '" & txtDRNum.Text & "' Order By DRItemNo ": CommandExecute
     Set mmsADORst = Nothing
     cmdCancel.Enabled = True
 End If

LocalError:
    Exit Sub
End Sub
Private Sub LoadDRDetails()
On Error GoTo LocalError
    lvwDR.ListItems.Clear
        strsql = "SELECT * from DRDetailsTemp ORDER BY DRItemNo ": CommandExecute
        With mmsADORst
        Do Until .EOF
        Set DRLI = lvwDR.ListItems.Add(, , !DRItemNo & "")
           DRLI.SubItems(1) = !DRHDate
           DRLI.SubItems(2) = !DRArea
           DRLI.SubItems(3) = !DRSpecie
           DRLI.SubItems(4) = !DRHill
           DRLI.SubItems(5) = !DRBolt
           DRLI.SubItems(6) = !DRSize
           DRLI.SubItems(7) = !DRPcs
           DRLI.SubItems(8) = !DRBdFt
           .MoveNext
        Loop
      End With
      
      If EncodeMode = "S" Then
         EncodeMode = "SF"
      End If
      
    cmdAdd.SetFocus
LocalError:
    Exit Sub
End Sub
Private Sub GetDRTotalAMount()
    strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetailsTemp ": CommandExecute
    DRTotalAmount = mmsADORst.Fields!Subtotal
End Sub
Private Sub InsertDRTotals()
On Error GoTo LocalError
        strsql = "Update DRDetailsTemp SET DRTotalPcs = '" & txtDRTotalPcs.Text & "'"
        strsql = strsql & ", DRTotalBdFt = '" & txtDRTotalBdFt.Text & "'"
        strsql = strsql & ", DRCum = '" & txtDRTotalCuM.Text & "'"
        strsql = strsql & " where DRNum like '" & txtDRNum.Text & "'"
        CommandExecute
LocalError:
    Exit Sub
End Sub
Private Sub LoadDRItemEdit()
On Error GoTo LocalError
    strsql = "SELECT * FROM DRDetailsTemp Where DRItemNo like '" & ItemEdit & "'"
    CommandExecute
    DRDetails
    With mmsADORst.Fields
        Contracted = !DRContract
          If Contracted = True Then
             checkCOntracted.Value = 1
          Else
             checkCOntracted.Value = 0
          End If
        checkCOntracted_Click
        txtDRProduct.Text = !DRProductClass
        DRProduct = !DRProduct
        DRClass = !DRClass
        txtDRSackBox.Text = !DRSackBox
        txtDRQty.Text = !DRQty
        txtDRWeight.Text = !DRWeight
        txtDRCost.Text = !DRCost
        txtDRAmount.Text = !DRAMount
    End With
LocalError:
    Exit Sub
End Sub
Private Function GetNextDRID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(DRId) AS MaxID FROM DRDetails"
    Set mmsADORst = mmsAdoCmd.Execute
       If mmsADORst.EOF Then
           GetNextDRID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextDRID = 1
       Else
           GetNextDRID = mmsADORst!MaxID + 1
       End If
    Set mmsADORst = Nothing
End Function
Private Function DataValidation() As Boolean
DataValidation = False
    If Not IsDate(txtDRDate.Text) Then
        MsgBox "Invalid Date.", vbExclamation, "Invalid Entry"
        txtDRDate.SetFocus
        txtDRDate = Format$(Now, "mm/dd/yyyy")
    End If
    If txtDRDelivered.Text = "" Then
        MsgBox "Fill-up Information", vbExclamation, "Buyer Required"
        txtDRDelivered.SetFocus
        Exit Function
    End If
DataValidation = True
End Function
Private Sub BoxState(boxEnabled As Boolean)
    txtDRDate.Enabled = boxEnabled
    txtDRDelivered.Enabled = boxEnabled
    txtDRDestination.Enabled = boxEnabled
    txtDRRemarks.Enabled = boxEnabled
End Sub
Private Sub ButtonState(buttonEnabled As Boolean)
    lvwDR.Enabled = buttonEnabled
    cmdNew.Enabled = buttonEnabled
    cmdSearch.Enabled = buttonEnabled
    cmdAdd.Enabled = buttonEnabled
    cmdCancel.Enabled = buttonEnabled
    cmdPrint.Enabled = buttonEnabled
End Sub
'------------ F O C U S ---------------
Private Sub cmdNew_GotFocus()
   cmdNew.BackColor = &HC0FFC0
End Sub
Private Sub cmdNew_LostFocus()
   cmdNew.BackColor = &H8000000F
End Sub
Private Sub cmdAdd_GotFocus()
   cmdAdd.BackColor = &HC0FFC0
End Sub
Private Sub cmdAdd_LostFocus()
   cmdAdd.BackColor = &H8000000F
End Sub
Private Sub cmdCancel_GotFocus()
   cmdCancel.BackColor = &HC0FFC0
End Sub
Private Sub cmdCancel_LostFocus()
   cmdCancel.BackColor = &H8000000F
End Sub
Private Sub cmdSearch_GotFocus()
   cmdSearch.BackColor = &HC0FFC0
End Sub
Private Sub cmdSearch_LostFocus()
   cmdSearch.BackColor = &H8000000F
End Sub
Private Sub cmdPrint_GotFocus()
   cmdPrint.BackColor = &HC0FFC0
End Sub
Private Sub cmdPrint_LostFocus()
   cmdPrint.BackColor = &H8000000F
End Sub
Private Sub cmdReport_GotFocus()
   cmdReport.BackColor = &HC0FFC0
End Sub
Private Sub cmdReport_LostFocus()
   cmdReport.BackColor = &H8000000F
End Sub
Private Sub cmdSaveDRDetails_GotFocus()
   cmdSaveDRDetails.BackColor = &HC0FFFF
End Sub
Private Sub cmdSaveDRDetails_LostFocus()
   cmdSaveDRDetails.BackColor = &H8000000F
End Sub
Private Sub cmdExit_GotFocus()
   cmdExit.BackColor = &HC0FFC0
End Sub
Private Sub cmdExit_LostFocus()
   cmdExit.BackColor = &H8000000F
End Sub
Private Sub SendKeys_()
    'SendKeys "{left}"
    'SendKeys "{del}"
End Sub

