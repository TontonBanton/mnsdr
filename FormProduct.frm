VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormProduct 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PRODUCT"
   ClientHeight    =   10725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9195
   Icon            =   "FormProduct.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10725
   ScaleMode       =   0  'User
   ScaleWidth      =   15668.28
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtProduct 
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
      ItemData        =   "FormProduct.frx":1AC52
      Left            =   1320
      List            =   "FormProduct.frx":1AC74
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   8200
      Width           =   4577
   End
   Begin VB.TextBox txtEncodeMode 
      Alignment       =   1  'Right Justify
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
      Height          =   450
      Left            =   8160
      MaxLength       =   50
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboUnit 
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
      ItemData        =   "FormProduct.frx":1ACD3
      Left            =   7035
      List            =   "FormProduct.frx":1ACDD
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   8760
      Width           =   1935
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7042
      MaxLength       =   50
      TabIndex        =   6
      Top             =   8200
      Width           =   1937
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
      Height          =   420
      Left            =   120
      MaxLength       =   50
      TabIndex        =   15
      Top             =   7440
      Visible         =   0   'False
      Width           =   1050
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
      Left            =   5135
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9600
      Width           =   1643
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
      Left            =   147
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9600
      Width           =   1643
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
      Left            =   1819
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9600
      Width           =   1643
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
      Left            =   7365
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9600
      Width           =   1643
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9600
      Width           =   1643
   End
   Begin VB.TextBox txtProduct1 
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
      Height          =   450
      Left            =   120
      MaxLength       =   50
      TabIndex        =   4
      Top             =   6960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtClassification 
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
      Height          =   450
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   5
      Top             =   8760
      Width           =   4577
   End
   Begin MSComctlLib.ListView lvwProduct 
      Height          =   7095
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   9210
      _ExtentX        =   16245
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
   Begin VB.Label Label3 
      BackColor       =   &H80000008&
      BorderStyle     =   1  'Fixed Single
      Height          =   1020
      Left            =   0
      TabIndex        =   17
      Top             =   9450
      Width           =   9184
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRICE"
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
      Index           =   3
      Left            =   6240
      TabIndex        =   16
      Top             =   8300
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "UNIT"
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
      Index           =   1
      Left            =   6240
      TabIndex        =   13
      Top             =   8835
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRODUCT"
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
      Index           =   2
      Left            =   150
      TabIndex        =   8
      Top             =   8300
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "CLASS"
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
      Index           =   0
      Left            =   150
      TabIndex        =   7
      Top             =   8805
      Width           =   660
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
      Height          =   600
      Left            =   120
      Picture         =   "FormProduct.frx":1ACEE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   825
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9184
   End
End
Attribute VB_Name = "FormProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn             As ADODB.Connection
Private mmsAdoCmd              As ADODB.Command
Private mmsADORst              As ADODB.Recordset
Private strsql                 As String
Private EncodeMode, Price      As String
Private ButtonPress            As String
Private Search                 As Boolean
Private Sub Form_Load()
   Load Me
   ConnectToDB
   SetlvwProduct
   LoadProduct
   lblComp.Caption = FormMainMenu.lblComp.Caption
   lblHeader.Caption = FormMainMenu.lblHeader.Caption
End Sub
Private Sub txtProduct_GotFocus()
   'txtProduct.SelStart = 0
   'txtProduct.SelLength = Len(txtProduct.Text)
End Sub
Private Sub txtProduct_KeyPress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
     If EncodeMode = "A" Then
        txtClassification.SetFocus
     Else
        txtClassification.SetFocus
     End If
  End If
End Sub
Private Sub txtClassification_GotFocus()
   txtClassification.SelLength = Len(txtClassification.Text)
End Sub
Private Sub txtClassification_Keypress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
     If EncodeMode = "A" Then
        txtPrice.SetFocus
     Else
        txtPrice.SetFocus
     End If
  End If
End Sub
Private Sub txtPrice_GotFocus()
   txtPrice.SelLength = Len(txtPrice.Text)
End Sub
Private Sub txtPrice_Keypress(KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = 46 Or KeyAscii = 13 Then
        If KeyAscii = 13 Then
            txtPrice.Text = Format$(txtPrice.Text, "#,###.#0")
            cboUnit.SetFocus
        End If
        If Chr$(KeyAscii) = "." And InStr(txtPrice, ".") > 0 Then
            KeyAscii = 0
        End If
        
      Else
       KeyAscii = 0
    End If
End Sub
Private Sub cboUnit_KeyPress(KeyAscii As Integer)
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
    txtProduct.SetFocus
End Sub
Private Sub CmdSave_Click()
    Dim mmsNewListItem  As ListItem
    Dim lngIDField      As Long
    Dim strsql          As String
 
On Error GoTo LocalError

   If Not DataValidation Then
      Exit Sub
   End If
    Price = Format$(txtPrice.Text, "#,###.#0")
    
      If EncodeMode = "A" Then

         lngIDField = GetNextProductID
    
         strsql = "INSERT INTO Product (  ProductID"
         strsql = strsql & "            , Product"
         strsql = strsql & "            , Classification"
         strsql = strsql & "            , Unit"
         strsql = strsql & "            , Price"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & lngIDField
         strsql = strsql & ", '" & Replace$(txtProduct.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtClassification.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(cboUnit.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(Price, "'", "''") & "'"
         strsql = strsql & ")"
             
          
          Set mmsNewListItem = lvwProduct.ListItems.Add(, , txtProduct.Text)
            PopulateItem mmsNewListItem
              With mmsNewListItem
                .SubItems(4) = CStr(lngIDField)
                .EnsureVisible
              End With
          Set lvwProduct.SelectedItem = mmsNewListItem
          Set mmsNewListItem = Nothing
          
      ElseIf EncodeMode = "U" Or EncodeMode = "S" Then 'Update
         lngIDField = lvwProduct.SelectedItem.SubItems(4)
         
         strsql = "UPDATE Product SET "
         strsql = strsql & "  Product         = '" & Replace$(txtProduct.Text, "'", "''") & "'"
         strsql = strsql & ", Classification  = '" & Replace$(txtClassification.Text, "'", "''") & "'"
         strsql = strsql & ", Unit            = '" & Replace$(cboUnit.Text, "'", "''") & "'"
         strsql = strsql & ", Price           = '" & Replace$(Price, "'", "''") & "'"
         strsql = strsql & " WHERE ProductID = " & lngIDField
          
         lvwProduct.SelectedItem.Text = txtProduct.Text
         PopulateItem lvwProduct.SelectedItem
    End If
    
    
    ConnectToDB
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    BoxState False
    lvwProduct.Enabled = True
    lvwProduct.SetFocus
    ButtonState True
    LoadProduct
    
LocalError:
    Exit Sub
End Sub
Private Sub cmdCancel_Click()
   Form_Load
   BoxState False
   ButtonState True
   lvwProduct.Enabled = True
   
    If lvwProduct.SelectedItem Is Nothing Then
        MsgBox "No Item selected to delete.", vbExclamation, "Delete"
        ClearBox
        Exit Sub
    End If
End Sub
Private Sub cmdDelete_Click()
    strsql = "Delete From Product where ProductID like " & txtID.Text & ""
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    MsgBox "Item Deleted"
   
   Form_Load
   BoxState False
   ButtonState True
   lvwProduct.Enabled = True
   
    If lvwProduct.SelectedItem Is Nothing Then
        MsgBox "No Item selected to delete.", vbExclamation, "Delete"
        ClearBox
        Exit Sub
    End If
End Sub
Private Sub cmdExit_Click()

    'If txtEncodeMode.Text = "PRODUCT" Then
    '     Unload Me
    '     FormMainMenu.Show
    'ElseIf txtEncodeMode.Text = "PRODUCT2" Then
    '     Unload Me
    '     FormDR.Enabled = True
    '     FormDR.Show
    'Else
    '    Exit Sub
    'End If
    
 Unload Me
 FormMainMenu.Show
End Sub
Private Sub Form_Unload(Cancel As Integer)
  cmdExit_Click
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
Private Sub lvwProduct_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo LocalError
     With Item
        txtProduct.Text = .Text
        txtClassification.Text = .SubItems(1)
        txtPrice.Text = .SubItems(2)
        cboUnit.Text = .SubItems(3)
        txtID.Text = .SubItems(4)
     End With
    'ButtonState True
   ' cmdCancel.Enabled = True
LocalError:
   Exit Sub
End Sub
Private Sub lvwProduct_DblClick()
    ButtonState False
    BoxState True
    txtProduct.Enabled = False
    txtClassification.SetFocus
    If EncodeMode = "S" Then
           With lvwProduct
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwProduct_ItemClick .SelectedItem
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
Private Sub SetlvwProduct()
    With lvwProduct
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Product", .Width * 0.35
        .ColumnHeaders.Add , , "Classification", .Width * 0.35
        .ColumnHeaders.Add , , "Price", .Width * 0.15
        .ColumnHeaders.Add , , "Unit", .Width * 0.1
        .ColumnHeaders.Add , , "ID", Width * 0#
    End With
  lvwProduct.ColumnHeaders.Item(3).Alignment = lvwColumnRight
End Sub
Private Sub LoadProduct()
    Dim strsql       As String
                          
    strsql = "SELECT * From Product Order By ProductID"
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    LVProduct
              With lvwProduct
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwProduct_ItemClick .SelectedItem
                  End If
              End With
End Sub
Private Sub LVProduct()
Dim ProductLI  As ListItem
lvwProduct.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set ProductLI = lvwProduct.ListItems.Add(, , !Product & "")
                ProductLI.SubItems(1) = !Classification
                ProductLI.SubItems(2) = !Price
                ProductLI.SubItems(3) = !Unit
                ProductLI.SubItems(4) = !ProductID
                .MoveNext
        Loop
     End With
    Set ProductLI = Nothing
    Set mmsADORst = Nothing
End Sub
Private Function DataValidation() As Boolean
  DataValidation = False
    'If txtProduct.Text = "" Then
    '     MsgBox "Fill-up Item Product.", vbExclamation, " ProductRequired"
    '     txtProduct.SetFocus
    '    Exit Function
    'End If
    If txtClassification.Text = "" Then
         MsgBox "Fill-up Item Classification.", vbExclamation, " ProductRequired"
         txtClassification.SetFocus
        Exit Function
    End If
    'If txtUnit.Text = "" Then
    '     MsgBox "Fill-up Item Unit.", vbExclamation, " ProductRequired"
    '     txtUnit.SetFocus
    '    Exit Function
    'End If
        If txtPrice.Text = "" Then
         MsgBox "Fill-up Item Price.", vbExclamation, " ProductRequired"
         txtPrice.SetFocus
        Exit Function
    End If
    If EncodeMode = "U" Then
    Else
    If txtProduct.Text = "" Then
        MsgBox "Fill-up Product", vbExclamation, "ProductRequired"
        txtProduct.SetFocus
        Exit Function
    End If
    End If
    DataValidation = True

End Function
Private Function GetNextProductID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(ProductID) AS MaxID FROM Product"
    Set mmsADORst = mmsAdoCmd.Execute
       
       If mmsADORst.EOF Then
           GetNextProductID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextProductID = 1
       Else
           GetNextProductID = mmsADORst!MaxID + 1
       End If

    Set mmsADORst = Nothing

End Function
Private Sub PopulateItem(mmsListItem As ListItem)
    With mmsListItem
        .SubItems(1) = txtClassification.Text
        .SubItems(2) = txtPrice.Text
        .SubItems(3) = cboUnit.Text
    End With
End Sub
Private Sub ClearBox()
    'txtProduct.Text = ""
    txtClassification.Text = ""
    txtPrice.Text = ""
    'txtUnit.Text = ""
End Sub
Private Sub BoxState(boxEnabled As Boolean)
    txtProduct.Enabled = boxEnabled
    txtClassification.Enabled = boxEnabled
    'txtUnit.Enabled = boxEnabled
    txtPrice.Enabled = boxEnabled
    cboUnit.Enabled = boxEnabled
End Sub
Private Sub ButtonState(buttonEnabled As Boolean)
    lvwProduct.Enabled = buttonEnabled
    cmdAdd.Enabled = buttonEnabled
    cmdExit.Enabled = buttonEnabled
    cmdCancel.Enabled = Not buttonEnabled
    cmdDelete.Enabled = Not buttonEnabled
    cmdSave.Enabled = Not buttonEnabled
End Sub
Private Sub SendKeys_()
    'SendKeys "{left}"
    'SendKeys "{del}"
End Sub

