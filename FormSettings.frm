VERSION 5.00
Begin VB.Form FormSettings 
   BackColor       =   &H80000002&
   Caption         =   "SETTINGS"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   Icon            =   "FormSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Height          =   4270
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   5595
      Begin VB.TextBox txtAreaLoc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2640
         TabIndex        =   30
         Top             =   360
         Width           =   2625
      End
      Begin VB.CommandButton cmdSettings 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3620
         Width           =   5025
      End
      Begin VB.TextBox txtChecked 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   2040
         TabIndex        =   4
         Top             =   1820
         Width           =   3225
      End
      Begin VB.TextBox txtNoted 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   2040
         TabIndex        =   7
         Top             =   3000
         Width           =   3225
      End
      Begin VB.TextBox txtAudited 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   2040
         TabIndex        =   6
         Top             =   2600
         Width           =   3225
      End
      Begin VB.TextBox txtApproved 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   2040
         TabIndex        =   5
         Top             =   2200
         Width           =   3225
      End
      Begin VB.TextBox txtPrepared 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   2040
         TabIndex        =   3
         Top             =   1440
         Width           =   3225
      End
      Begin VB.TextBox txtLocation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   3705
      End
      Begin VB.TextBox txtArea 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Height          =   620
         Left            =   120
         TabIndex        =   28
         Top             =   3540
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AUDITED BY"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   500
         TabIndex        =   23
         Top             =   2650
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CHECKED BY"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   500
         TabIndex        =   22
         Top             =   1920
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "APPROVED BY"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   500
         TabIndex        =   21
         Top             =   2300
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NOTED BY"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   500
         TabIndex        =   19
         Top             =   3020
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PREPARED BY"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   500
         TabIndex        =   18
         Top             =   1530
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LOCATION"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Top             =   750
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AREA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   15
         Top             =   400
         Width           =   435
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Left            =   120
         TabIndex        =   17
         Top             =   200
         Width           =   5295
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   2145
         Left            =   120
         TabIndex        =   20
         Top             =   1305
         Width           =   5295
      End
   End
   Begin VB.TextBox txtComp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   200
      Width           =   5595
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   2760
      TabIndex        =   14
      Top             =   6900
      Width           =   2950
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   2685
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000002&
      Height          =   2115
      Left            =   120
      TabIndex        =   24
      Top             =   4750
      Width           =   5600
      Begin VB.CommandButton cmdUser 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1400
         Width           =   5025
      End
      Begin VB.TextBox txtUser 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   3225
      End
      Begin VB.TextBox txtPWord 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   720
         Width           =   3225
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Height          =   620
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "USER NAME"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   360
         TabIndex        =   26
         Top             =   450
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   360
         TabIndex        =   25
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   5295
      End
   End
End
Attribute VB_Name = "FormSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mmsADOConn             As ADODB.Connection
Private mmsAdoCmd              As ADODB.Command
Private mmsADORst              As ADODB.Recordset
Private strsql                 As String
Private Sub Form_Load()
   ConnectToDB
   LoadSettings
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Unload Me
     FormMainMenu.cmdExit.SetFocus
End Sub

'-----------------------------------------------------------------
'------------------------------------------------------------------
Private Sub txtComp_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtArea.SetFocus
   End If
End Sub
Private Sub txtArea_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtAreaLoc.SetFocus
   End If
End Sub
Private Sub txtAreaLoc_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtLocation.SetFocus
   End If
End Sub
Private Sub txtLocation_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtPrepared.SetFocus
   End If
End Sub
Private Sub txtPrepared_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtChecked.SetFocus
   End If
End Sub
Private Sub txtChecked_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtApproved.SetFocus
   End If
End Sub
Private Sub txtApproved_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtAudited.SetFocus
   End If
End Sub
Private Sub txtAudited_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtNoted.SetFocus
   End If
End Sub
Private Sub txtNoted_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      cmdSettings.SetFocus
   End If
End Sub
Private Sub txtUser_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtPWord.SetFocus
   End If
End Sub
Private Sub txtPWord_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      cmdUser.SetFocus
   End If
End Sub
'----------------------------------------------------------------
Private Sub txtComp_GotFocus()
  txtComp.SelLength = Len(txtComp.Text)
End Sub
Private Sub txtArea_DblClick()
  txtArea.Text = ""
End Sub
Private Sub txtArea_GotFocus()
  txtArea.SelLength = Len(txtArea.Text)
End Sub
Private Sub txtAreaLoc_DblClick()
  txtAreaLoc.Text = ""
End Sub
Private Sub txtAreaLoc_GotFocus()
  txtAreaLoc.SelLength = Len(txtAreaLoc.Text)
End Sub
Private Sub txtLocation_DblClick()
  txtLocation.Text = ""
End Sub
Private Sub txtLocation_GotFocus()
  txtLocation.SelLength = Len(txtLocation.Text)
End Sub
Private Sub txtPrepared_DblClick()
  txtPrepared.Text = ""
End Sub
Private Sub txtPrepared_GotFocus()
  txtPrepared.SelLength = Len(txtPrepared.Text)
End Sub
Private Sub txtChecked_DblClick()
  txtChecked.Text = ""
End Sub
Private Sub txtChecked_GotFocus()
  txtChecked.SelLength = Len(txtChecked.Text)
End Sub
Private Sub txtApproved_DblClick()
  txtApproved.Text = ""
End Sub
Private Sub txtApproved_GotFocus()
  txtApproved.SelLength = Len(txtApproved.Text)
End Sub
Private Sub txtAudited_DblClick()
  txtAudited.Text = ""
End Sub
Private Sub txtAudited_GotFocus()
  txtAudited.SelLength = Len(txtAudited.Text)
End Sub
Private Sub txtNoted_DblClick()
  txtNoted.Text = ""
End Sub
Private Sub txtNoted_GotFocus()
  txtNoted.SelLength = Len(txtNoted.Text)
End Sub
Private Sub txtUser_DblClick()
  txtNoted.Text = ""
End Sub
Private Sub txtPWord_DblClick()
  txtReceived.Text = ""
End Sub


'---------------------------------------------------------------------

'--------------------------------------------------------------------
Private Sub cmdSettings_Click()
    If MsgBox("Save current settings?", _
        vbYesNo + vbQuestion, "Settings") = vbYes Then
        strsql = "Update Settings SET Area = '" & txtArea.Text & "'"
        strsql = strsql & ", AreaLoc = '" & txtAreaLoc.Text & "'"
        strsql = strsql & ", Location = '" & txtLocation.Text & "'"
        strsql = strsql & ", Company = '" & txtComp.Text & "'"
        strsql = strsql & ", Prepared = '" & txtPrepared.Text & "'"
        strsql = strsql & ", Checked = '" & txtChecked.Text & "'"
        strsql = strsql & ", Approved = '" & txtApproved.Text & "'"
        strsql = strsql & ", Audited = '" & txtAudited.Text & "'"
        strsql = strsql & ", Noted = '" & txtNoted.Text & "'"
        mmsAdoCmd.CommandText = strsql
        Set mmsADORst = mmsAdoCmd.Execute
        cmdClose.SetFocus
    Else
        LoadSettings
        cmdClose.SetFocus
    End If
End Sub
Private Sub cmdUser_Click()
Dim UserId As Double
         UserId = 1
         strsql = "INSERT INTO PWord ( UName"
         strsql = strsql & " , PWord"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & " '" & Replace$(txtUser.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtPWord.Text, "'", "''") & "'"
         strsql = strsql & ")"
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    Set mmsADORst = Nothing
    txtUser.Text = ""
    txtPWord.Text = ""
    cmdClose.SetFocus
End Sub

Private Sub cmdClose_Click()
   If MsgBox("Are you sure you want to exit?", _
        vbYesNo + vbQuestion, "Exit") = vbYes Then
     'ClearBox
     'ClearItemBox
     'ClearFrame
     Unload Me
     FormMainMenu.Show
     FormMainMenu.cmdExit.SetFocus
   Else
     Exit Sub
 End If
End Sub
'--------------------------------------------------------------------

'-------------------------------------------------------------------
Private Sub ConnectToDB()
    Set mmsADOConn = New ADODB.Connection
    mmsADOConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MMS2.mdb" & ";Persist Security Info=False"
    mmsADOConn.Open
    
        mstrSQL = "select * from DRDetails"
        Set mmsADORst = New ADODB.Recordset
        mmsADORst.CursorLocation = adUseClient
        mmsADORst.Open mstrSQL, mmsADOConn, adOpenDynamic, adLockOptimistic, adCmdText
      

    Set mmsAdoCmd = New ADODB.Command
    Set mmsAdoCmd.ActiveConnection = mmsADOConn
    mmsAdoCmd.CommandType = adCmdText

End Sub
Private Sub LoadSettings()
On Error GoTo LocalError
   mmsAdoCmd.CommandText = "Select * from Settings"
   Set mmsADORst = mmsAdoCmd.Execute
   txtComp.Text = mmsADORst.Fields("Company")
   txtArea.Text = mmsADORst.Fields("Area")
   txtAreaLoc.Text = mmsADORst.Fields("AreaLoc")
   txtLocation.Text = mmsADORst.Fields("Location")
   txtPrepared.Text = mmsADORst.Fields("Prepared")
   txtApproved.Text = mmsADORst.Fields("Approved")
   txtAudited.Text = mmsADORst.Fields("Audited")
   txtNoted.Text = mmsADORst.Fields("Noted")
   txtChecked.Text = mmsADORst.Fields("Checked")
LocalError:
    Exit Sub
End Sub

Private Sub cmdSettings_GotFocus()
   cmdSettings.BackColor = &HC0FFC0
End Sub
Private Sub cmdSettings_LostFocus()
   cmdSettings.BackColor = &H8000000F
End Sub
Private Sub cmdUser_GotFocus()
   cmdUser.BackColor = &HC0FFC0
End Sub
Private Sub cmdUser_LostFocus()
   cmdUser.BackColor = &H8000000F
End Sub
Private Sub cmdClose_GotFocus()
   cmdClose.BackColor = &HC0FFC0
End Sub
Private Sub cmdClose_LostFocus()
   cmdClose.BackColor = &H8000000F
End Sub
