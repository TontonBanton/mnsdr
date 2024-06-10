VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormDestination 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BUYER"
   ClientHeight    =   9780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDestination 
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
      Height          =   420
      Left            =   150
      MaxLength       =   50
      TabIndex        =   7
      Top             =   8100
      Width           =   6650
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
      Height          =   650
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8930
      Width           =   1500
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
      Height          =   650
      Left            =   1650
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8930
      Width           =   1500
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
      Height          =   650
      Left            =   5300
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8930
      Width           =   1500
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
      Height          =   650
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8930
      Width           =   1500
   End
   Begin MSComctlLib.ListView lvwDestination 
      Height          =   7095
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   12515
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
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      BorderStyle     =   1  'Fixed Single
      Height          =   1020
      Left            =   0
      TabIndex        =   9
      Top             =   8760
      Width           =   7000
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
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   120
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "FormDueTo.frx":0000
      Stretch         =   -1  'True
      Top             =   50
      Width           =   780
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   1020
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7000
   End
End
Attribute VB_Name = "FormDestination"
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
   SetlvwDestination
   LoadDestination
   lblComp.Caption = FormMainMenu.lblComp.Caption
   lblHeader.Caption = FormMainMenu.lblHeader.Caption
End Sub
Private Sub txtDestination_GotFocus()
   txtDestination.SelLength = Len(txtDestination.Text)
End Sub
Private Sub txtDestination_Keypress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
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
    txtDestination.SetFocus
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

         lngIDField = GetNextDestinationID
    
         strsql = "INSERT INTO Destination (  DestinationID"
         strsql = strsql & "            , Destination"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & lngIDField
         strsql = strsql & ", '" & Replace$(txtDestination.Text, "'", "''") & "'"
         strsql = strsql & ")"
             
          
          Set mmsNewListItem = lvwDestination.ListItems.Add(, , txtDestination.Text)
            PopulateItem mmsNewListItem
              With mmsNewListItem
                .SubItems(1) = CStr(lngIDField)
                .EnsureVisible
              End With
          Set lvwDestination.SelectedItem = mmsNewListItem
          Set mmsNewListItem = Nothing
          
      ElseIf EncodeMode = "U" Or EncodeMode = "S" Then 'Update
         lngIDField = lvwDestination.SelectedItem.SubItems(1)
         
         strsql = "UPDATE Destination SET "
         strsql = strsql & "  Destination         = '" & Replace$(txtDestination.Text, "'", "''") & "'"
         strsql = strsql & " WHERE DestinationID = " & lngIDField
          
         lvwDestination.SelectedItem.Text = txtDestination.Text
         PopulateItem lvwDestination.SelectedItem
    End If
    ConnectToDB
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    BoxState False
    lvwDestination.Enabled = True
    lvwDestination.SetFocus
    ButtonState True
    

LocalError:
    Exit Sub
End Sub
Private Sub cmdCancel_Click()
   Form_Load
   BoxState False
   ButtonState True
   lvwDestination.Enabled = True
   
    If lvwDestination.SelectedItem Is Nothing Then
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
Private Sub cmdExit_GotFocus()
   cmdExit.BackColor = &HC0FFC0
End Sub
Private Sub cmdExit_LostFocus()
   cmdExit.BackColor = &H8000000F
End Sub

'--------------------------------------------------------------------------
'                             L I S T V I E W
'--------------------------------------------------------------------------
Private Sub lvwDestination_ItemClick(ByVal Item As MSComctlLib.ListItem)
     With Item
        txtDestination.Text = .Text
     End With
    'ButtonState True
   ' cmdCancel.Enabled = True
End Sub
Private Sub lvwDestination_DblClick()
    ButtonState False
    BoxState True
    txtDestination.SetFocus
    If EncodeMode = "S" Then
           With lvwDestination
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwDestination_ItemClick .SelectedItem
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
Private Sub SetlvwDestination()
    With lvwDestination
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Destination", .Width * 0.98
        .ColumnHeaders.Add , , "ID", Width * 0#
    End With
End Sub
Private Sub LoadDestination()
    Dim strsql       As String

                                 
    strsql = "SELECT * From Destination Order By Destination"
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    LVDestination
              With lvwDestination
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwDestination_ItemClick .SelectedItem
                  End If
              End With
End Sub
Private Sub LVDestination()
Dim DestinationLI  As ListItem
lvwDestination.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set DestinationLI = lvwDestination.ListItems.Add(, , !Destination & "")
                DestinationLI.SubItems(1) = !DestinationID
                .MoveNext
        Loop
     End With
    Set DestinationLI = Nothing
    Set mmsADORst = Nothing
End Sub
Private Function DataValidation() As Boolean
  DataValidation = False
    If txtDestination.Text = "" Then
         MsgBox "Fill-up Item Destination.", vbExclamation, " DestinationRequired"
         txtDestination.SetFocus
        Exit Function
    End If
    If EncodeMode = "U" Then
    Else
    If txtDestination.Text = "" Then
        MsgBox "Fill-up Destination", vbExclamation, "DestinationRequired"
        txtDestination.SetFocus
        Exit Function
    End If
    End If
    DataValidation = True

End Function
Private Function GetNextDestinationID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(DestinationID) AS MaxID FROM Destination"
    Set mmsADORst = mmsAdoCmd.Execute
       
       If mmsADORst.EOF Then
           GetNextDestinationID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextDestinationID = 1
       Else
           GetNextDestinationID = mmsADORst!MaxID + 1
       End If

    Set mmsADORst = Nothing

End Function
Private Sub PopulateItem(mmsListItem As ListItem)
    'With mmsListItem
    '    .SubItems(1) = txtDestination.Text
    'End With
End Sub
Private Sub ClearBox()
    txtDestination.Text = ""
End Sub
Private Sub BoxState(boxEnabled As Boolean)
    txtDestination.Enabled = boxEnabled
End Sub
Private Sub ButtonState(buttonEnabled As Boolean)
    lvwDestination.Enabled = buttonEnabled
    cmdAdd.Enabled = buttonEnabled
    cmdExit.Enabled = buttonEnabled
    cmdCancel.Enabled = Not buttonEnabled
    'cmdDelete.Enabled = Not buttonEnabled
    cmdSave.Enabled = Not buttonEnabled
End Sub


