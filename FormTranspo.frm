VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormTranspo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSPORTATION"
   ClientHeight    =   10725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10725
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVehID 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6960
      MaxLength       =   50
      TabIndex        =   13
      Top             =   7320
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.TextBox txtPlate 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   10
      Top             =   8760
      Width           =   6300
   End
   Begin VB.TextBox txtVehicle 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   0
      Top             =   8200
      Width           =   6300
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4650
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9650
      Width           =   1550
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CANCEL"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3100
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9650
      Width           =   1550
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9650
      Width           =   1600
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9650
      Width           =   1500
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   50
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9650
      Width           =   1550
   End
   Begin MSComctlLib.ListView lvwTranspo 
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   12303
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483646
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
   Begin VB.Label Label3 
      BackColor       =   &H80000008&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   0
      TabIndex        =   14
      Top             =   9480
      Width           =   8050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "VEHICLE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   300
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   8250
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "PLATE NO."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   8800
      Width           =   1050
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   960
      TabIndex        =   9
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
      Left            =   960
      TabIndex        =   8
      Top             =   120
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   120
      Picture         =   "FormTranspo.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   850
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8000
   End
End
Attribute VB_Name = "FormTranspo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn             As ADODB.Connection
Private mmsAdoCmd              As ADODB.Command
Private mmsADORst              As ADODB.Recordset
Private strsql                 As String
Private EncodeMode             As String
Private ButtonPress            As String
Private Search                 As Boolean
Private Sub Form_Load()
   Load Me
   ConnectToDB
   SetlvwTranspo
   LoadTranspo
   lblComp.Caption = FormMainMenu.lblComp.Caption
   lblHeader.Caption = FormMainMenu.lblHeader.Caption
End Sub
Private Sub txtVehicle_GotFocus()
   txtVehicle.SelLength = Len(txtVehicle.Text)
End Sub
Private Sub txtVehicle_Keypress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
     If EncodeMode = "A" Then
        txtPlate.SetFocus
     Else
        txtPlate.SetFocus
     End If
  End If
End Sub
Private Sub txtPlate_GotFocus()
   txtPlate.SelLength = Len(txtPlate.Text)
End Sub
Private Sub txtPlate_Keypress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
     'txtTranspo.Text = txtVehicle.Text & " " & txtPlate.Text
     If EncodeMode = "A" Then
        cmdSave.SetFocus
     Else
        cmdSave.SetFocus
     End If
  End If
End Sub
Private Sub CmdAdd_Click()
    EncodeMode = "A"
    ButtonState False
    BoxState True
    ClearBox
    txtVehicle.SetFocus
End Sub
Private Sub CmdSave_Click()
    Dim mmsNewListItem  As ListItem
    Dim lngIDField      As Long
    Dim strsql          As String
 
On Error GoTo LocalError

   If Not DataValidation Then
      Exit Sub
   End If
      If EncodeMode = "A" Then

         lngIDField = GetNextVehicleID
         
         strsql = "INSERT INTO Vehicle (  VehicleID"
         strsql = strsql & "            , PlateNo"
         strsql = strsql & "            , VehicleName"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & lngIDField
         strsql = strsql & ", '" & Replace$(txtPlate.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtVehicle.Text, "'", "''") & "'"
         strsql = strsql & ")"
             
          
          Set mmsNewListItem = lvwTranspo.ListItems.Add(, , txtVehicle.Text)
            PopulateItem mmsNewListItem
              With mmsNewListItem
                .SubItems(2) = CStr(lngIDField)
                .EnsureVisible
              End With
          Set lvwTranspo.SelectedItem = mmsNewListItem
          Set mmsNewListItem = Nothing
          
      ElseIf EncodeMode = "U" Or EncodeMode = "S" Then 'Update
         lngIDField = lvwTranspo.SelectedItem.SubItems(2)
         
         strsql = "UPDATE Vehicle SET "
         strsql = strsql & "   PlateNo        = '" & Replace$(txtPlate.Text, "'", "''") & "'"
         strsql = strsql & ",  VehicleName    = '" & Replace$(txtVehicle.Text, "'", "''") & "'"
         strsql = strsql & " WHERE VehicleID  = " & lngIDField
         
         lvwTranspo.SelectedItem.Text = txtVehicle.Text
         PopulateItem lvwTranspo.SelectedItem
    End If
    
    ConnectToDB
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    BoxState False
    lvwTranspo.Enabled = True
    lvwTranspo.SetFocus
    ButtonState True
    
LocalError:
    Exit Sub
End Sub
Private Sub cmdCancel_Click()
   Form_Load
   BoxState False
   ButtonState True
   lvwTranspo.Enabled = True
   
    If lvwTranspo.SelectedItem Is Nothing Then
        MsgBox "No Item selected to delete.", vbExclamation, "Delete"
        ClearBox
        Exit Sub
    End If
End Sub
Private Sub cmdDelete_Click()
    strsql = "Delete From Vehicle where VehicleID like " & txtVehID.Text & ""
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    MsgBox "Item Deleted"
   
   Form_Load
   BoxState False
   ButtonState True
   lvwTranspo.Enabled = True
   
    If lvwTranspo.SelectedItem Is Nothing Then
        MsgBox "No Item selected to delete.", vbExclamation, "Delete"
        ClearBox
        Exit Sub
    End If
End Sub
Private Sub cmdExit_Click()
     Unload Me
End Sub
'------------ F O C U S ---------------
Private Sub cmdAdd_GotFocus()
   cmdAdd.BackColor = &HC0FFC0
End Sub
Private Sub cmdAdd_LostFocus()
   cmdAdd.BackColor = &H8000000F
End Sub
Private Sub cmdSave_GotFocus()
   cmdSave.BackColor = &HC0FFC0
End Sub
Private Sub cmdSave_LostFocus()
   cmdSave.BackColor = &H8000000F
End Sub
Private Sub cmdCancel_GotFocus()
   cmdCancel.BackColor = &HC0FFC0
End Sub
Private Sub cmdCancel_LostFocus()
   cmdCancel.BackColor = &H8000000F
End Sub
Private Sub cmdDelete_GotFocus()
   cmdDelete.BackColor = &HC0FFC0
End Sub
Private Sub cmdDelete_LostFocus()
   cmdDelete.BackColor = &H8000000F
End Sub
Private Sub cmdExit_GotFocus()
   cmdExit.BackColor = &HC0FFC0
End Sub
Private Sub cmdExit_LostFocus()
   cmdExit.BackColor = &H8000000F
End Sub
'--------------------------------------------------------------------------
'                             L I S T V I E W
'--------------------------------------------------------------------------
Private Sub lvwTranspo_ItemClick(ByVal Item As MSComctlLib.ListItem)
     With Item
        txtVehicle.Text = .Text
        txtPlate.Text = .SubItems(1)
        txtVehID.Text = .SubItems(2)
     End With
    'ButtonState True
   ' cmdCancel.Enabled = True
End Sub
Private Sub lvwTranspo_DblClick()
    ButtonState False
    BoxState True
    txtVehicle.SetFocus
    If EncodeMode = "S" Then
           With lvwTranspo
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwTranspo_ItemClick .SelectedItem
                  End If
           End With
    Else
            EncodeMode = "U"
    End If
End Sub
'--------------------------------------------------------------------------
'                                 F U N C T I O N S
'--------------------------------------------------------------------------
Private Sub ConnectToDB()
    Set mmsADOConn = New ADODB.Connection
    mmsADOConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MMS2.mdb" & ";Persist Security Info=False"
    mmsADOConn.Open
    
        strsql = "select * from DRDetails"
        Set mmsADORst = New ADODB.Recordset
        mmsADORst.CursorLocation = adUseClient
        mmsADORst.Open strsql, mmsADOConn, adOpenDynamic, adLockPessimistic, adCmdText
      

    Set mmsAdoCmd = New ADODB.Command
    Set mmsAdoCmd.ActiveConnection = mmsADOConn
    mmsAdoCmd.CommandType = adCmdText

End Sub
Private Sub SetlvwTranspo()
    With lvwTranspo
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "", .Width * 0.5
        .ColumnHeaders.Add , , "", .Width * 0.2
        .ColumnHeaders.Add , , "ID", Width * 0
    End With
End Sub
Private Sub LoadTranspo()
    Dim strsql       As String
                          
    strsql = "SELECT * From Vehicle Order By VehicleName"
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    LVTranspo
              With lvwTranspo
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwTranspo_ItemClick .SelectedItem
                  End If
              End With
End Sub
Private Sub LVTranspo()
Dim TranspoLI  As ListItem
lvwTranspo.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set TranspoLI = lvwTranspo.ListItems.Add(, , !VehicleName & "")
                TranspoLI.SubItems(1) = !PlateNo
                TranspoLI.SubItems(2) = !VehicleID
                .MoveNext
        Loop
     End With
    Set TranspoLI = Nothing
    Set mmsADORst = Nothing
End Sub
Private Function DataValidation() As Boolean
  DataValidation = False
    If txtVehicle.Text = "" Then
         MsgBox "Fill-up Vehicle Model.", vbExclamation, " ModelRequired"
         txtVehicle.SetFocus
        Exit Function
    End If
    If txtPlate.Text = "" Then
         MsgBox "Fill-up Item Classification.", vbExclamation, " PlateNoRequired"
         txtPlate.SetFocus
        Exit Function
    End If
    If EncodeMode = "U" Then
    Else
    If txtVehicle.Text = "" Then
        MsgBox "Fill-up Model", vbExclamation, "ModelRequired"
        txtVehicle.SetFocus
        Exit Function
    End If
    End If
    DataValidation = True

End Function
Private Function GetNextVehicleID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(VehicleID) AS MaxID FROM Vehicle"
    Set mmsADORst = mmsAdoCmd.Execute
       
       If mmsADORst.EOF Then
           GetNextVehicleID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextVehicleID = 1
       Else
           GetNextVehicleID = mmsADORst!MaxID + 1
       End If

    Set mmsADORst = Nothing

End Function
Private Sub PopulateItem(mmsListItem As ListItem)
    With mmsListItem
        .SubItems(1) = txtPlate.Text
        .SubItems(2) = txtVehID.Text
    End With
End Sub
Private Sub ClearBox()
    txtVehicle.Text = ""
    txtPlate.Text = ""
    txtVehID.Text = ""
End Sub
Private Sub BoxState(boxEnabled As Boolean)
    txtVehicle.Enabled = boxEnabled
    txtPlate.Enabled = boxEnabled
    txtVehID.Enabled = boxEnabled
End Sub
Private Sub ButtonState(buttonEnabled As Boolean)
    lvwTranspo.Enabled = buttonEnabled
    cmdAdd.Enabled = buttonEnabled
    cmdExit.Enabled = buttonEnabled
    cmdCancel.Enabled = Not buttonEnabled
    cmdDelete.Enabled = Not buttonEnabled
    cmdSave.Enabled = Not buttonEnabled
End Sub
