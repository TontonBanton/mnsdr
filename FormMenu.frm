VERSION 5.00
Begin VB.Form FormMainMenu 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MS COMPANY INCORPORATED  "
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   16905
   ForeColor       =   &H80000008&
   Icon            =   "FormMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   16905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInv 
      Caption         =   "INVENTORY"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   960
      Width           =   2500
   End
   Begin VB.CommandButton cmdDR 
      Caption         =   "DELIVERY RECEIPT"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   2500
   End
   Begin VB.Frame frameTools 
      BackColor       =   &H80000001&
      Height          =   2535
      Left            =   8160
      TabIndex        =   17
      Top             =   1680
      Visible         =   0   'False
      Width           =   2500
      Begin VB.CommandButton cmdSettings 
         Caption         =   "SETTINGS"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1660
         Width           =   2000
      End
      Begin VB.CommandButton cmdConvert 
         Caption         =   "CONVERSION TOOLS"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   2000
      End
      Begin VB.CommandButton cmdCalculator 
         Caption         =   "CALCULATOR"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1000
         Width           =   2000
      End
   End
   Begin VB.CommandButton cmdTools 
      Caption         =   "&TOOLS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   2500
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   14280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.Frame frameLib 
      BackColor       =   &H80000001&
      Height          =   3975
      Left            =   5640
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   2500
      Begin VB.CommandButton cmdArea 
         Caption         =   "AREA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2970
         Width           =   2000
      End
      Begin VB.CommandButton cmdTranspo 
         Caption         =   "TRANSPORT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2320
         Width           =   2000
      End
      Begin VB.CommandButton cmdDestination 
         Caption         =   "DESTINATION"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1000
         Width           =   2000
      End
      Begin VB.CommandButton cmdProduct 
         Caption         =   "PRODUCT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1660
         Width           =   2000
      End
      Begin VB.CommandButton cmdDelivered 
         Caption         =   "BUYER"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   350
         Width           =   2000
      End
   End
   Begin VB.CommandButton cmdLib 
      Caption         =   "&LIBRARIES"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   2500
   End
   Begin VB.Timer Timer1 
      Left            =   13800
      Top             =   360
   End
   Begin VB.Image Image5 
      Height          =   5745
      Left            =   2400
      Picture         =   "FormMenu.frx":1E72
      Stretch         =   -1  'True
      Top             =   2925
      Width           =   11880
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USER :"
      Height          =   195
      Left            =   1080
      TabIndex        =   22
      Top             =   720
      Width           =   540
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AREA"
      Height          =   195
      Left            =   1680
      TabIndex        =   21
      Top             =   720
      Width           =   435
   End
   Begin VB.Label lblArea 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AREA"
      Height          =   195
      Left            =   2400
      TabIndex        =   20
      Top             =   150
      Visible         =   0   'False
      Width           =   435
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
      Height          =   240
      Left            =   1080
      TabIndex        =   19
      Top             =   150
      Width           =   1125
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "Location - Adress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   18
      Top             =   400
      Width           =   1560
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   800
      Left            =   0
      TabIndex        =   11
      Top             =   900
      Width           =   16920
   End
   Begin VB.Label lblClock 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Label3"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   14760
      TabIndex        =   14
      Top             =   555
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   15960
      Picture         =   "FormMenu.frx":31C63
      Stretch         =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblToday 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Label3"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   14400
      TabIndex        =   13
      Top             =   315
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   870
      Left            =   0
      Picture         =   "FormMenu.frx":37D1E
      Stretch         =   -1  'True
      Top             =   50
      Width           =   945
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   14280
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image Image4 
      Height          =   7245
      Left            =   2160
      Picture         =   "FormMenu.frx":3CA8E
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   12285
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Height          =   945
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   16995
   End
   Begin VB.Menu mnuForm 
      Caption         =   "Transaction &Forms"
      Index           =   1
      Begin VB.Menu mnuDR 
         Caption         =   "Delivery Receipt"
      End
   End
   Begin VB.Menu mnuForm 
      Caption         =   "&Libraries"
      Index           =   2
      Begin VB.Menu mnuDelivered 
         Caption         =   "Delivered"
      End
      Begin VB.Menu mnuDestination 
         Caption         =   "Destination"
      End
      Begin VB.Menu mnuProduct 
         Caption         =   "Product"
      End
      Begin VB.Menu mnuTranspo 
         Caption         =   "Transportation"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Index           =   3
      Begin VB.Menu mnuConvet 
         Caption         =   "Conversion"
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
   End
End
Attribute VB_Name = "FormMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn             As ADODB.Connection
Private mmsAdoCmd              As ADODB.Command
Private mmsADORst              As ADODB.Recordset
Private strsql                 As String
Private Sub Form_Activate()
On Error GoTo LocalError
  ConnectToDB
  mmsAdoCmd.CommandText = "Select * From Settings"
  Set mmsADORst = mmsAdoCmd.Execute
  lblComp.Caption = mmsADORst.Fields("Company")
  lblHeader.Caption = mmsADORst.Fields("Area") & "-" & mmsADORst.Fields("Location")
  lblArea.Caption = mmsADORst.Fields("Area")
  lblUser.Caption = FormLog.UserName
LocalError:
  Exit Sub
End Sub
Private Sub Form_Load()
  Load Me
  Timer1.Interval = 100
  frameLib.Visible = False
  ConnectToDB
  Form_Activate
End Sub
Private Sub mnuDelivered_Click()
   cmdExit.SetFocus
   FormDelivered.Show
End Sub
Private Sub mnuDestination_Click()
   cmdExit.SetFocus
   FormDestination.Show
End Sub
Private Sub mnuArea_Click()
   cmdExit.SetFocus
   FormArea.Show
End Sub
Private Sub mnuProduct_Click()
   cmdExit.SetFocus
   'FormLog.txtEncodeMode.Text = "PRODUCT"
   'FormLog.Show
   FormProduct.Show
End Sub
Private Sub mnuTranspo_Click()
   cmdExit.SetFocus
   FormTranspo.Show
End Sub
Private Sub mnuAdjust_Click()
   cmdExit.SetFocus
   FormLog.txtEncodeMode.Text = "ADJUST"
   FormLog.Show
End Sub
Private Sub mnuDR_Click()
   cmdDR.SetFocus
   FormDR.Show
   Unload Me
End Sub
Private Sub mnuConvert_Click()
   cmdExit.SetFocus
   'FormConvert.Show
   Shell "explorer.exe http://www.unitconverters.net/"
   
End Sub
Private Sub mnuSettings_Click()
   cmdExit.SetFocus
   FormLog.txtEncodeMode.Text = "SETTINGS"
   FormLog.Show
End Sub
Private Sub mnuCalculator_Click()
   cmdExit.SetFocus
  Shell "C:\Windows\System32\calc.exe"
End Sub
Private Sub mnuExit_Click()
  If MsgBox("Are you sure you want to exit?", _
       vbYesNo + vbQuestion, "Exit") = vbYes Then
       End
  Else
       Exit Sub
  End If
End Sub
Private Sub Timer1_Timer()
   lblToday.Caption = Format$(Now, "ddd, mmm/dd/yyyy")
   lblClock.Caption = Format$(Now, "hh:mm AM/PM")
End Sub
'------------------------------------------------------------------------
'                      B U T T O N S   E V E N T S
'-------------------------------------------------------------------------
Private Sub cmdLib_Click()
  frameLib.Visible = True
  frameTools.Visible = False
  cmdDelivered.SetFocus
End Sub
Private Sub CmdTools_Click()
  frameTools.Visible = True
  frameLib.Visible = False
  cmdConvert.SetFocus
End Sub
Private Sub cmdExit_Click()
  mnuExit_Click
End Sub
Private Sub CmdAdjust_Click()
  mnuAdjust_Click
End Sub
Private Sub cmdAdjust_KeyPress(KeyAscii As Integer)
  cmdLib_GotFocus
End Sub
Private Sub CmdDR_Click()
  mnuDR_Click
End Sub
Private Sub cmdDR_KeyPress(KeyAscii As Integer)
  cmdLib_GotFocus
End Sub
Private Sub CmdDelivered_Click()
  mnuDelivered_Click
End Sub
Private Sub CmdDestination_Click()
  mnuDestination_Click
End Sub
Private Sub CmdProduct_Click()
  mnuProduct_Click
End Sub
Private Sub CmdTranspo_Click()
  mnuTranspo_Click
End Sub
Private Sub CmdArea_Click()
  mnuArea_Click
End Sub
Private Sub cmdArea_KeyPress(KeyAscii As Integer)
  cmdTools_GotFocus
End Sub
Private Sub cmdConvert_Click()
  'MsgBox "Work On Progress", vbInformation, "Convert"
  mnuConvert_Click
End Sub
Private Sub cmdConvert_KeyPress(KeyAscii As Integer)
  cmdExit_GotFocus
End Sub
Private Sub cmdCalculator_Click()
  mnuCalculator_Click
End Sub
Private Sub cmdCalculator_KeyPress(KeyAscii As Integer)
  cmdExit_GotFocus
End Sub
Private Sub cmdSettings_Click()
  mnuSettings_Click
End Sub
Private Sub cmdSettings_KeyPress(KeyAscii As Integer)
  cmdExit_GotFocus
End Sub
'------------ F O C U S ---------------
Private Sub cmdLib_GotFocus()
   cmdLib.BackColor = &HC0FFC0
   frameLib.Visible = True
   frameTools.Visible = False
End Sub
Private Sub cmdLib_LostFocus()
   cmdLib.BackColor = &H8000000F
End Sub
Private Sub cmdTools_GotFocus()
   cmdTools.BackColor = &HC0FFC0
   frameTools.Visible = True
   frameLib.Visible = False
End Sub
Private Sub cmdTools_LostFocus()
   cmdTools.BackColor = &H8000000F
End Sub
Private Sub cmdExit_GotFocus()
   cmdExit.BackColor = &HC0FFC0
   frameLib.Visible = False
   frameTools.Visible = False
End Sub
Private Sub cmdExit_LostFocus()
   cmdExit.BackColor = &H8000000F
End Sub
Private Sub cmdDelivered_GotFocus()
   cmdDelivered.BackColor = &HC0FFC0
End Sub
Private Sub cmdDelivered_LostFocus()
   cmdDelivered.BackColor = &H8000000F
End Sub
Private Sub cmdDestination_GotFocus()
   cmdDestination.BackColor = &HC0FFC0
End Sub
Private Sub cmdDestination_LostFocus()
   cmdDestination.BackColor = &H8000000F
End Sub
Private Sub cmdProduct_GotFocus()
   cmdProduct.BackColor = &HC0FFC0
End Sub
Private Sub cmdProduct_LostFocus()
   cmdProduct.BackColor = &H8000000F
End Sub
Private Sub cmdArea_GotFocus()
   cmdArea.BackColor = &HC0FFC0
End Sub
Private Sub cmdArea_LostFocus()
   cmdArea.BackColor = &H8000000F
End Sub
Private Sub cmdDR_GotFocus()
   frameLib.Visible = False
   frameTools.Visible = False
   cmdDR.BackColor = &HC0FFC0
End Sub
Private Sub cmdDR_LostFocus()
   cmdDR.BackColor = &H8000000F
End Sub
Private Sub cmdConvert_GotFocus()
   cmdConvert.BackColor = &HC0FFC0
End Sub
Private Sub cmdConvert_LostFocus()
   cmdConvert.BackColor = &H8000000F
End Sub
Private Sub cmdCalculator_GotFocus()
   cmdCalculator.BackColor = &HC0FFC0
End Sub
Private Sub cmdCalculator_LostFocus()
   cmdCalculator.BackColor = &H8000000F
End Sub
Private Sub cmdSettings_GotFocus()
   cmdSettings.BackColor = &HC0FFC0
End Sub
Private Sub cmdSettings_LostFocus()
   cmdSettings.BackColor = &H8000000F
End Sub
Private Sub ConnectToDB()
    Set mmsADOConn = New ADODB.Connection
    mmsADOConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MMS2.mdb" & ";Persist Security Info=False"
    mmsADOConn.Open
    
        strsql = "Select * from DRDetails"
        Set mmsADORst = New ADODB.Recordset
        mmsADORst.CursorLocation = adUseClient
        mmsADORst.Open strsql, mmsADOConn, adOpenDynamic, adLockPessimistic, adCmdText
      

    Set mmsAdoCmd = New ADODB.Command
    Set mmsAdoCmd.ActiveConnection = mmsADOConn
    mmsAdoCmd.CommandType = adCmdText

End Sub


