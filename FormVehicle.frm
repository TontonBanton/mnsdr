VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormDelivered 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DESTINATION"
   ClientHeight    =   10725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10725
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
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
      Height          =   650
      Left            =   5000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9480
      Width           =   1600
   End
   Begin VB.ComboBox txtSaleTo 
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
      ItemData        =   "FormVehicle.frx":0000
      Left            =   2040
      List            =   "FormVehicle.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   8160
      Width           =   4335
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
      Left            =   3370
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9480
      Width           =   1600
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9480
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
      Height          =   650
      Left            =   1740
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9480
      Width           =   1600
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9480
      Width           =   1600
   End
   Begin VB.TextBox txtCustomer 
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
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   6
      Top             =   8760
      Width           =   4290
   End
   Begin VB.TextBox txtID 
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
      Left            =   6840
      MaxLength       =   50
      TabIndex        =   4
      Top             =   8160
      Visible         =   0   'False
      Width           =   1650
   End
   Begin MSComctlLib.ListView lvwDelivered 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   12091
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648384
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
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   850
      Left            =   -360
      TabIndex        =   13
      Top             =   9360
      Width           =   9500
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
         Height          =   461
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1550
      End
      Begin VB.CommandButton Command1 
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
         Height          =   461
         Left            =   13320
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1550
      End
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
         Height          =   461
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1550
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "CUSTOMER"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   300
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "SALE TO"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   300
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   8280
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   240
      Picture         =   "FormVehicle.frx":0059
      Stretch         =   -1  'True
      Top             =   240
      Width           =   780
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
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   1125
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
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   690
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   780
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8685
   End
End
Attribute VB_Name = "FormDelivered"
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

Private Const M_Customer                  As Long = 1
Private Const M_DeliveredID               As Long = 2
Private Sub Form_Load()
   Load Me
   ConnectToDB
   SetlvwDelivered
   LoadDelivered
   lblComp.Caption = FormMainMenu.lblComp.Caption
   lblHeader.Caption = FormMainMenu.lblHeader.Caption
End Sub
Private Sub txtSaleTo_GotFocus()
   'txtSaleTo.SelStart = 0
   'txtSaleTo.SelLength = Len(txtSaleTo.Text)
End Sub
Private Sub txtSaleTo_Keypress(KeyAscii As Integer)
  'KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
     If EncodeMode = "A" Then
        txtCustomer.SetFocus
     Else
        txtCustomer.SetFocus
     End If
  End If
End Sub
Private Sub txtCustomer_GotFocus()
   txtCustomer.SelLength = Len(txtCustomer.Text)
End Sub
Private Sub txtcustomer_Keypress(KeyAscii As Integer)
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
    txtSaleTo.SetFocus
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

         lngIDField = GetNextDeliveredID
    
         strsql = "INSERT INTO Delivered (  DeliveredID"
         strsql = strsql & "            , SaleTo"
         strsql = strsql & "            , Customer"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & lngIDField
         strsql = strsql & ", '" & Replace$(txtSaleTo.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtCustomer.Text, "'", "''") & "'"
         strsql = strsql & ")"
             
          
          Set mmsNewListItem = lvwDelivered.ListItems.Add(, , txtSaleTo.Text)
            PopulateItem mmsNewListItem
              With mmsNewListItem
                .SubItems(M_DeliveredID) = CStr(lngIDField)
                .EnsureVisible
              End With
          Set lvwDelivered.SelectedItem = mmsNewListItem
          Set mmsNewListItem = Nothing
          
      ElseIf EncodeMode = "U" Or EncodeMode = "S" Then 'Update
         lngIDField = lvwDelivered.SelectedItem.SubItems(M_DeliveredID)
         
         strsql = "UPDATE Delivered SET "
         strsql = strsql & "  SaleTo         = '" & Replace$(txtSaleTo.Text, "'", "''") & "'"
         strsql = strsql & ", Customer            = '" & Replace$(txtCustomer.Text, "'", "''") & "'"
         strsql = strsql & " WHERE DeliveredID = " & lngIDField
          
         lvwDelivered.SelectedItem.Text = txtSaleTo.Text
         PopulateItem lvwDelivered.SelectedItem
    End If
    
    ConnectToDB
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    BoxState False
    lvwDelivered.Enabled = True
    lvwDelivered.SetFocus
    ButtonState True
    

LocalError:
    Exit Sub
End Sub
Private Sub cmdCancel_Click()
   Form_Load
   BoxState False
   ButtonState True
   lvwDelivered.Enabled = True
   
    If lvwDelivered.SelectedItem Is Nothing Then
        MsgBox "No Item selected to delete.", vbExclamation, "Delete"
        ClearBox
        Exit Sub
    End If
End Sub
Private Sub cmdDelete_Click()
        strsql = "Delete From Delivered where DeliveredID like " & txtID.Text & ""
        mmsAdoCmd.CommandText = strsql
        Set mmsADORst = mmsAdoCmd.Execute
        MsgBox "Item Deleted"
   Form_Load
   BoxState False
   ButtonState True
   lvwDelivered.Enabled = True
   
    If lvwDelivered.SelectedItem Is Nothing Then
        MsgBox "No Item selected to delete.", vbExclamation, "Delete"
        ClearBox
        Exit Sub
    End If
End Sub
Private Sub CmdExit_Click()
   If MsgBox("Are you sure you want to exit?", _
        vbYesNo + vbQuestion, "Exit") = vbYes Then
     Unload Me
 Else
     Exit Sub
 End If
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
Private Sub lvwDelivered_ItemClick(ByVal Item As MSComctlLib.ListItem)
     With Item
        txtSaleTo.Text = .Text
        txtCustomer.Text = .SubItems(M_Customer)
        txtID.Text = .SubItems(2)
     End With
    'ButtonState True
   ' cmdCancel.Enabled = True
End Sub
Private Sub lvwDelivered_DblClick()
    ButtonState False
    BoxState True
    txtSaleTo.SetFocus
    If EncodeMode = "S" Then
           With lvwDelivered
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwDelivered_ItemClick .SelectedItem
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
    
        strsql = "select * from DRDetailsTemp"
        Set mmsADORst = New ADODB.Recordset
        mmsADORst.CursorLocation = adUseClient
        mmsADORst.Open strsql, mmsADOConn, adOpenDynamic, adLockPessimistic, adCmdText
      

    Set mmsAdoCmd = New ADODB.Command
    Set mmsAdoCmd.ActiveConnection = mmsADOConn
    mmsAdoCmd.CommandType = adCmdText

End Sub
Private Sub SetlvwDelivered()
    With lvwDelivered
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "SaleTo", .Width * 0.48
        .ColumnHeaders.Add , , "Customer", .Width * 0.48
        .ColumnHeaders.Add , , "ID", Width * 0#
    End With
End Sub
Private Sub LoadDelivered()
    Dim strsql       As String
                            
    strsql = "SELECT * From Delivered Order By SaleTo, Customer"
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    LVDelivered
              With lvwDelivered
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwDelivered_ItemClick .SelectedItem
                  End If
              End With
End Sub
Private Sub LVDelivered()
Dim DeliveredLI  As ListItem
lvwDelivered.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set DeliveredLI = lvwDelivered.ListItems.Add(, , !SaleTo & "")
                DeliveredLI.SubItems(M_Customer) = !customer
                DeliveredLI.SubItems(M_DeliveredID) = !DeliveredID
                .MoveNext
        Loop
     End With
    Set DeliveredLI = Nothing
    Set mmsADORst = Nothing
End Sub
Private Function DataValidation() As Boolean
  DataValidation = False
    If txtSaleTo.Text = "" Then
         MsgBox "Fill-up Item SaleTo.", vbExclamation, " SaleToRequired"
         txtSaleTo.SetFocus
        Exit Function
    End If
    If EncodeMode = "U" Then
    Else
    If txtSaleTo.Text = "" Then
        MsgBox "Fill-up SaleTo", vbExclamation, "SaleToRequired"
        txtSaleTo.SetFocus
        Exit Function
    End If
    End If
    DataValidation = True

End Function
Private Function GetNextDeliveredID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(DeliveredID) AS MaxID FROM Delivered"
    Set mmsADORst = mmsAdoCmd.Execute
       
       If mmsADORst.EOF Then
           GetNextDeliveredID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextDeliveredID = 1
       Else
           GetNextDeliveredID = mmsADORst!MaxID + 1
       End If

    Set mmsADORst = Nothing

End Function
Private Sub PopulateItem(mmsListItem As ListItem)
    With mmsListItem
        .SubItems(M_Customer) = txtCustomer.Text
    End With
End Sub
Private Sub ClearBox()
   ' txtSaleTo.Text = ""
    txtCustomer.Text = ""
End Sub
Private Sub BoxState(boxEnabled As Boolean)
    txtSaleTo.Enabled = boxEnabled
    txtCustomer.Enabled = boxEnabled
End Sub
Private Sub ButtonState(buttonEnabled As Boolean)
    lvwDelivered.Enabled = buttonEnabled
    cmdAdd.Enabled = buttonEnabled
    cmdExit.Enabled = buttonEnabled
    cmdCancel.Enabled = Not buttonEnabled
    cmdDelete.Enabled = Not buttonEnabled
    cmdSave.Enabled = Not buttonEnabled
End Sub

