VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormInv 
   BackColor       =   &H80000002&
   Caption         =   "MS COMPANY INCORPORATED  "
   ClientHeight    =   10545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16755
   Icon            =   "FormInv.frx":0000
   LinkTopic       =   "FormInv"
   MaxButton       =   0   'False
   ScaleHeight     =   10545
   ScaleMode       =   0  'User
   ScaleWidth      =   16755
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtInvSize 
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
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1920
      Width           =   2700
   End
   Begin VB.ComboBox txtInvClass 
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
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1440
      Width           =   2700
   End
   Begin VB.ComboBox txtInvArea 
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
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   960
      Width           =   2700
   End
   Begin MSComCtl2.DTPicker txtinvDate 
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   550
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   108134401
      CurrentDate     =   45454
   End
   Begin VB.TextBox txtInvPcs 
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
      Height          =   700
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   9
      Top             =   4320
      Width           =   2700
   End
   Begin VB.TextBox txtInvBolt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   5
      Top             =   3600
      Width           =   2715
   End
   Begin VB.TextBox txtInvHill 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   3
      Top             =   2880
      Width           =   2700
   End
   Begin MSComctlLib.ListView lvwInv 
      Height          =   10560
      Left            =   4440
      TabIndex        =   0
      Top             =   0
      Width           =   12330
      _ExtentX        =   21749
      _ExtentY        =   18627
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "PCS"
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
      Left            =   200
      TabIndex        =   10
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "AREA"
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
      Left            =   200
      TabIndex        =   8
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "BOLT #"
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
      Index           =   5
      Left            =   200
      TabIndex        =   7
      Top             =   3600
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "SIZE"
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
      Index           =   3
      Left            =   200
      TabIndex        =   6
      Top             =   2040
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "HILL #"
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
      Index           =   1
      Left            =   200
      TabIndex        =   4
      Top             =   2880
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "SPECIE"
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
      Index           =   0
      Left            =   200
      TabIndex        =   2
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "DATE CUT"
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
      Left            =   200
      TabIndex        =   1
      Top             =   600
      Width           =   945
   End
End
Attribute VB_Name = "FormInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn             As ADODB.Connection
Private mmsAdoCmd              As ADODB.Command
Private mmsADORst              As ADODB.Recordset
Public strsql, EncodeMode      As String
Private DRLI       As ListItem
Private Sub Form_Load()
   EncodeMode = "INVENTORY"
   ConnectToDB
   LoadInventory
   txtinvDate.Value = Now
   LoadArea
   LoadClass
   LoadSizes
   'lblComp.Caption = FormMainMenu.lblComp.Caption
   'lblHeader.Caption = FormMainMenu.lblHeader.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Unload Me
     FormMainMenu.Show
End Sub
Private Sub LoadInventory()
SetInventory
strsql = "SELECT * from Inventory ORDER By InvID": CommandExecute
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
Private Sub LoadArea()
    strsql = "SELECT * FROM Area ORDER BY Area": CommandExecute
    With mmsADORst
       txtInvArea.Clear
        Do While Not .EOF
            txtInvArea.AddItem ![Area]
            .MoveNext
        Loop
    End With
End Sub
Private Sub LoadClass()
    strsql = "SELECT Class FROM Product ORDER BY Class": CommandExecute
    With mmsADORst
       txtInvClass.Clear
        Do While Not .EOF
            txtInvClass.AddItem ![Class]
            .MoveNext
        Loop
    End With
End Sub
Private Sub LoadSizes()
    strsql = "SELECT SizeName FROM Sizes ORDER BY SizeName": CommandExecute
    With mmsADORst
        txtInvSize.Clear
        Do While Not .EOF
            txtInvSize.AddItem ![SizeName]
            .MoveNext
        Loop
    End With
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
Private Sub txtInvArea_Click()
    txtInvClass.SetFocus
End Sub
Private Sub txtInvClass_Click()
    txtInvSize.SetFocus
End Sub
Private Sub txtInvSize_Click()
    txtInvHill.SetFocus
End Sub
