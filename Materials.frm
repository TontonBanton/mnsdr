VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormItems 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "    ITEMS LIBRARY "
   ClientHeight    =   10725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "Materials.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10725
   ScaleMode       =   0  'User
   ScaleWidth      =   15270
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1995
      Left            =   11050
      TabIndex        =   27
      Top             =   7530
      Width           =   3960
      Begin VB.ComboBox cboItemSearch 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Materials.frx":08CA
         Left            =   480
         List            =   "Materials.frx":08D4
         TabIndex        =   11
         Top             =   840
         Width           =   3045
      End
      Begin VB.TextBox txtItemSearch 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1200
         Width           =   3045
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SEARCH OPTIONS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Index           =   5
         Left            =   1000
         TabIndex        =   34
         Top             =   480
         Width           =   2010
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Height          =   1635
         Left            =   150
         TabIndex        =   33
         Top             =   240
         Width           =   3645
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1995
      Left            =   280
      TabIndex        =   16
      Top             =   7530
      Width           =   10800
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   15
         Left            =   0
         TabIndex        =   26
         Top             =   2160
         Width           =   135
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   15
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   4335
      End
      Begin VB.TextBox txtAvailStock 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8400
         MaxLength       =   10
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1860
      End
      Begin VB.ComboBox CboGroup 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Materials.frx":08F0
         Left            =   2280
         List            =   "Materials.frx":0900
         TabIndex        =   8
         Top             =   960
         Width           =   2205
      End
      Begin VB.TextBox txtDes 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   7
         Top             =   480
         Width           =   4290
      End
      Begin VB.ComboBox cboUnit 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Materials.frx":092C
         Left            =   2280
         List            =   "Materials.frx":0939
         TabIndex        =   9
         Top             =   1320
         Width           =   2205
      End
      Begin VB.TextBox txtCost 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4680
         MaxLength       =   13
         TabIndex        =   10
         Top             =   1320
         Width           =   1905
      End
      Begin VB.TextBox txtMinStock 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8400
         MaxLength       =   10
         TabIndex        =   15
         Top             =   840
         Width           =   1860
      End
      Begin VB.TextBox TxtMCode 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4680
         MaxLength       =   15
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   960
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AVAILABLE "
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
         Index           =   6
         Left            =   7200
         TabIndex        =   21
         Top             =   1320
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "MINIMUM "
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
         Index           =   3
         Left            =   7200
         TabIndex        =   17
         Top             =   900
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "GROUP / CODE"
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
         Index           =   0
         Left            =   600
         TabIndex        =   20
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "UNIT / COST"
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
         Index           =   1
         Left            =   600
         TabIndex        =   19
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION"
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
         Index           =   2
         Left            =   600
         TabIndex        =   18
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Height          =   1635
         Left            =   150
         TabIndex        =   31
         Top             =   240
         Width           =   6730
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "STOCK INVENTORY "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Index           =   4
         Left            =   7800
         TabIndex        =   30
         Top             =   480
         Width           =   2205
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Height          =   1635
         Left            =   7000
         TabIndex        =   32
         Top             =   240
         Width           =   3645
      End
   End
   Begin MSComctlLib.ListView lvwMaterials 
      Height          =   6135
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   10821
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
      NumItems        =   0
   End
   Begin VB.Frame Frame2 
      Height          =   1155
      Left            =   280
      TabIndex        =   22
      Top             =   9400
      Width           =   14730
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&CANCEL"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1800
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&SAVE"
         Enabled         =   0   'False
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
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1800
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&XIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1800
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3260
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1800
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&UPDATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   5060
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1800
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&DELETE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   6860
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1800
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   8880
         TabIndex        =   29
         Top             =   240
         Width           =   5650
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   3120
         TabIndex        =   28
         Top             =   240
         Width           =   5650
      End
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
      Left            =   5280
      TabIndex        =   36
      Top             =   480
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
      Left            =   5280
      TabIndex        =   35
      Top             =   240
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   240
      Picture         =   "Materials.frx":094C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   4200
      Picture         =   "Materials.frx":1912D
      Stretch         =   -1  'True
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   15015
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   3270
      Left            =   120
      TabIndex        =   24
      Top             =   7410
      Width           =   15015
   End
End
Attribute VB_Name = "FormItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn             As ADODB.Connection
Private mmsAdoCmd              As ADODB.Command
Private mmsADORst              As ADODB.Recordset
Private mstrSQL                As String
Private EncodeMode             As String
Private ButtonPress            As String
Private Search                 As Boolean

Private Const M_Code                      As Long = 1
Private Const M_Group                     As Long = 2
Private Const M_Cost                      As Long = 3
Private Const M_Unit                      As Long = 4
Private Const M_MinStock                  As Long = 5
Private Const M_AvailStock                As Long = 6
Private Const M_ItemID                    As Long = 7

'-----------------------------------------------------------------------------------
'                                   F O R M   E V E N T S
'----------------------------------------------------------------------------------
Private Sub Form_Load()
   Load Me
   ConnectToDB
   SetlvwMaterials
   LoadMaterials
   cboItemSearch.Text = "DESCRIPTION"
   lblComp.Caption = FormMainMenu.lblComp.Caption
   lblHeader.Caption = FormMainMenu.lblHeader.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
     Unload Me
     FormMainMenu.cmdExit.SetFocus
End Sub

'---------------------------------------------------------------------------------
'                                   C O N T R O L S   E V E N T S
'---------------------------------------------------------------------------------
Private Sub txtDes_GotFocus()
   txtDes.SelStart = 0
   txtDes.SelLength = Len(txtDes.Text)
   cboItemSearch.Text = ""
End Sub
Private Sub TxtDes_KeyPress(KeyAscii As Integer)
  KeyAscii = ConvertUpper(KeyAscii)
  If KeyAscii = 13 Then
     If EncodeMode = "A" Then
        CboGroup.SetFocus
     Else
        cboUnit.SetFocus
     End If
  End If
End Sub
Private Sub CboGroup_KeyPress(KeyAscii As Integer)
      If KeyAscii > 0 Then
         If Chr$(KeyAscii) = "C" Or Chr$(KeyAscii) = "c" Then
           CboGroup.Text = "CHEMICALS"
           SendKeys_
          Else
          If Chr$(KeyAscii) = "F" Or Chr$(KeyAscii) = "f" Then
              CboGroup.Text = "FERTILIZERS"
              SendKeys_
            Else
            If Chr$(KeyAscii) = "M" Or Chr$(KeyAscii) = "m" Then
               CboGroup.Text = "MATERIALS"
               SendKeys_
               Else
                  If Chr$(KeyAscii) = "P" Or Chr$(KeyAscii) = "p" Then
                  CboGroup.Text = "POL"
                  SendKeys_
                  Else
                    If KeyAscii = 8 Then
                      Exit Sub
                    Else
                     If KeyAscii = 13 Then
                     cboUnit.SetFocus
                     Else
                        SendKeys_
                     End If
                  End If
               End If
             End If
           End If
          End If
      End If
      'ButtonShortcuts
End Sub
Private Sub CboGroup_LostFocus()
Dim ItemGroup As String
Dim ItemId    As String

If CboGroup.Text = "" Then
   Exit Sub
Else
  If EncodeMode = "A" Then
     ItemGroup = CboGroup.Text
     ItemId = Format$(GetNextItemID, "000000")
     TxtMCode.Text = Mid$(ItemGroup, 1, 1) & ItemId
  End If
End If
End Sub
Private Sub CboUnit_GotFocus()
If CboGroup.Text = "POL" Or CboGroup.Text = "CHEMICALS" Then
    cboUnit.Text = "LTS"
ElseIf CboGroup.Text = "MATERIALS" Then
    cboUnit.Text = "PCS"
ElseIf CboGroup.Text = "FERTILIZERS" Then
    cboUnit.Text = "KLO"
End If
End Sub
Private Sub CboUnit_Keypress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtCost.SetFocus
   ElseIf KeyAscii = 8 Then
      Exit Sub
   ElseIf IsNumeric(Chr(KeyAscii)) Then
      SendKeys_
   End If
End Sub
Private Sub txtCost_GotFocus()
  If EncodeMode = "A" Then
   txtCost.Text = ".00"
  End If
   txtCost.SelLength = Len(txtCost.Text)
End Sub
Private Sub txtCost_LostFocus()
Dim ItemCost  As Double
   ItemCost = Val(txtCost.Text)
   txtCost.Text = Format$(ItemCost, "###.#0")
End Sub
Private Sub TxtCost_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 And Chr$(KeyAscii) <> "." Then
      SendKeys_
   ElseIf KeyAscii = 13 Then
      txtMinStock.SetFocus
   End If
       If Chr$(KeyAscii) = "." And InStr(txtCost.Text, ".") > 0 Then
        KeyAscii = 0
       End If
End Sub
Private Sub txtMinStock_GotFocus()
  If EncodeMode = "A" Then
   txtMinStock.Text = 1
   txtAvailStock.Text = 0
  End If
  txtMinStock.SelLength = Len(txtMinStock.Text)
End Sub
Private Sub TxtMinStock_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 Then
      SendKeys_
   ElseIf KeyAscii = 13 Then
      cmdSave.SetFocus
   End If
End Sub
'------------------------------------------------------------------------------------
'                            S E A R C H   C O N T R O L S
'-----------------------------------------------------------------------------------
Private Sub cboItemSearch_GotFocus()
   'cboItemSearch.Text = "DESCRIPTION"
   cboItemSearch.SelStart = 0
   cboItemSearch.SelLength = Len(cboItemSearch.Text)
End Sub
Private Sub cboItemSearch_Click()
  'ButtonState False
  'cmdSave.Enabled = False
  lvwMaterials.Enabled = True
  Search = True
End Sub
Private Sub cboItemSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii < 255 Then
      SendKeys_
   End If
   If KeyAscii = 13 Then
      txtItemSearch.SetFocus
   End If
End Sub
Private Sub txtItemSearch_Change()
Dim strsql       As String
Dim MaterialsLI  As ListItem
On Error GoTo LocalError
    Search = True
    If cboItemSearch.Text = "DESCRIPTION" Then
       strsql = "Select * from Materials where ItemDescription like '" & txtItemSearch.Text & "%'" _
             & "Order by ItemDescription"
    Else
       strsql = "Select * from Materials where ItemCode like '" & txtItemSearch.Text & "%'" _
             & "Order by ItemCode"
    End If
      
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    LVMaterials
LocalError:
    Exit Sub
End Sub
Private Sub txtItemSearch_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
    lvwMaterials.SetFocus
   End If
End Sub
'--------------------------------------------------------------------------------------
'                        B U T T O N S   E V E N T S
'--------------------------------------------------------------------------------------
Private Sub CmdAdd_Click()
    EncodeMode = "A"
    ButtonState False
    BoxState True
    ClearBox
    txtDes.SetFocus
End Sub
Private Sub cmdAdd_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
   ButtonShortcuts
End Sub
Private Sub CmdUpdate_Click()
    ButtonState False
    BoxState True
    txtDes.SetFocus
    CboGroup.Enabled = False
    If EncodeMode = "S" Then
           With lvwMaterials
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwMaterials_ItemClick .SelectedItem
                  End If
           End With
    Else
            EncodeMode = "U"
    End If
End Sub
Private Sub cmdUpdate_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
   ButtonShortcuts
End Sub
Private Sub cmdDelete_Click()

    If lvwMaterials.SelectedItem Is Nothing Then
        MsgBox "No Item selected to delete.", vbExclamation, "Delete"
        Exit Sub
    End If

    If MsgBox("Are you sure that you want to delete the item on the list " _
              , vbYesNo + vbQuestion, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    ConnectToDB
    mmsAdoCmd.CommandText = "DELETE FROM Materials WHERE ItemCode =  '" & TxtMCode.Text & "'"
    mmsAdoCmd.Execute
    ClearBox
    LoadMaterials
End Sub
Private Sub cmdDelete_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
   ButtonShortcuts
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

         lngIDField = GetNextItemID
    
         strsql = "INSERT INTO Materials (  ItemID"
         strsql = strsql & "            , ItemGroup"
         strsql = strsql & "            , ItemCode"
         strsql = strsql & "            , ItemDescription"
         strsql = strsql & "            , Unit"
         strsql = strsql & "            , Cost"
         strsql = strsql & "            , MinStock"
         strsql = strsql & "            , AvailStock"
         strsql = strsql & "         ) VALUES ("
         strsql = strsql & lngIDField
         strsql = strsql & ", '" & Replace$(CboGroup.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(TxtMCode.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(txtDes.Text, "'", "''") & "'"
         strsql = strsql & ", '" & Replace$(cboUnit.Text, "'", "''") & "'"
         strsql = strsql & ", '" & txtCost.Text & "'"
         strsql = strsql & ", '" & txtMinStock.Text & "'"
         strsql = strsql & ", '" & txtAvailStock.Text & "'"
         strsql = strsql & ")"
             
          
          Set mmsNewListItem = lvwMaterials.ListItems.Add(, , txtDes.Text)
            PopulateItem mmsNewListItem
              With mmsNewListItem
                .SubItems(M_ItemID) = CStr(lngIDField)
                .EnsureVisible
              End With
          Set lvwMaterials.SelectedItem = mmsNewListItem
          Set mmsNewListItem = Nothing
          
      ElseIf EncodeMode = "U" Or EncodeMode = "S" Then 'Update
         lngIDField = CLng(lvwMaterials.SelectedItem.SubItems(M_ItemID))
         
         strsql = "UPDATE Materials SET "
         strsql = strsql & "  ItemGroup           = '" & Replace$(CboGroup.Text, "'", "''") & "'"
         strsql = strsql & ", Itemcode            = '" & Replace$(TxtMCode.Text, "'", "''") & "'"
         strsql = strsql & ", ItemDescription     = '" & Replace$(txtDes.Text, "'", "''") & "'"
         strsql = strsql & ", Unit                = '" & Replace$(cboUnit.Text, "'", "''") & "'"
         strsql = strsql & ", Cost                = '" & txtCost.Text & "'"
         strsql = strsql & ", MinStock            = '" & txtMinStock.Text & "'"
         strsql = strsql & ", AvailStock          = '" & txtAvailStock.Text & "'"
         strsql = strsql & " WHERE ItemID = " & lngIDField
          
         lvwMaterials.SelectedItem.Text = txtDes.Text
         PopulateItem lvwMaterials.SelectedItem
    End If
    
    ConnectToDB
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    BoxState False
    lvwMaterials.Enabled = True
    lvwMaterials.SetFocus
    ButtonState True
    
    cboItemSearch.Text = ""
    txtItemSearch.Text = ""

LocalError:
    Exit Sub
End Sub
Private Sub cmdCancel_Click()
   Form_Load
   BoxState False
   ButtonState True
   lvwMaterials.Enabled = True
   cboItemSearch.Text = ""
   txtItemSearch.Text = ""
    
    If lvwMaterials.SelectedItem Is Nothing Then
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
Private Sub cmdExit_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
   ButtonShortcuts
End Sub
'------------ F O C U S ---------------
Private Sub cmdAdd_GotFocus()
   cmdAdd.BackColor = &HC0FFC0
End Sub
Private Sub cmdAdd_LostFocus()
   cmdAdd.BackColor = &H8000000F
End Sub
Private Sub cmdUpdate_GotFocus()
   cmdUpdate.BackColor = &HC0FFC0
End Sub
Private Sub cmdUpdate_LostFocus()
   cmdUpdate.BackColor = &H8000000F
End Sub
Private Sub cmdDelete_GotFocus()
   cmdDelete.BackColor = &HC0FFC0
End Sub
Private Sub cmdDelete_LostFocus()
   cmdDelete.BackColor = &H8000000F
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
'                     L I S T V I E W    E V E N T S
'-------------------------------------------------------------------------
Private Sub lvwMaterials_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'LoadMaterials
    With lvwMaterials
        If (.Sorted) And (ColumnHeader.SubItemIndex = .SortKey) Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .Sorted = True
            .SortKey = ColumnHeader.SubItemIndex
            .SortOrder = lvwAscending
        End If
        .Refresh
    End With
        
    If Not lvwMaterials.SelectedItem Is Nothing Then
        lvwMaterials.SelectedItem.EnsureVisible
    End If

End Sub
Private Sub lvwMaterials_ItemClick(ByVal Item As MSComctlLib.ListItem)
     With Item
        txtDes.Text = .Text
        TxtMCode.Text = .SubItems(M_Code)
        CboGroup.Text = .SubItems(M_Group)
        cboUnit.Text = .SubItems(M_Unit)
        txtCost.Text = .SubItems(M_Cost)
        txtMinStock.Text = .SubItems(M_MinStock)
        txtAvailStock.Text = .SubItems(M_AvailStock)
     End With
    'ButtonState True
   ' cmdCancel.Enabled = True
End Sub
Private Sub lvwMaterials_KeyPress(KeyAscii As Integer)
   ButtonPress = Chr$(KeyAscii)
   ButtonShortcuts
   If KeyAscii = 13 Then
    cmdAdd.SetFocus
   End If
End Sub

'--------------------------------------------------------------------------
'                                 F U N C T I O N S
'--------------------------------------------------------------------------
Private Sub ConnectToDB()
    Set mmsADOConn = New ADODB.Connection
    mmsADOConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MMS2.mdb" & ";Persist Security Info=False"
    mmsADOConn.Open
    
        mstrSQL = "select * from Materials"
        Set mmsADORst = New ADODB.Recordset
        mmsADORst.CursorLocation = adUseClient
        mmsADORst.Open mstrSQL, mmsADOConn, adOpenDynamic, adLockOptimistic, adCmdText
      

    Set mmsAdoCmd = New ADODB.Command
    Set mmsAdoCmd.ActiveConnection = mmsADOConn
    mmsAdoCmd.CommandType = adCmdText

End Sub
Private Sub SetlvwMaterials()
    With lvwMaterials
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Description", .Width * 0.37
        .ColumnHeaders.Add , , "Code", .Width * 0.1
        .ColumnHeaders.Add , , "Group", .Width * 0.13
        .ColumnHeaders.Add , , "Cost", .Width * 0.12
        .ColumnHeaders.Add , , "Unit", .Width * 0.06
        .ColumnHeaders.Add , , "MinStock", .Width * 0.1
        .ColumnHeaders.Add , , "AvailStock", .Width * 0.1
        .ColumnHeaders.Add , , "ID", Width * 0#
    End With
lvwMaterials.ColumnHeaders.Item(4).Alignment = lvwColumnRight
End Sub
Private Sub PopulateItem(mmsListItem As ListItem)
    With mmsListItem
        .SubItems(M_Code) = TxtMCode.Text
        .SubItems(M_Unit) = cboUnit.Text
        .SubItems(M_Cost) = txtCost.Text
        .SubItems(M_Group) = CboGroup.Text
        .SubItems(M_MinStock) = txtMinStock.Text
        .SubItems(M_AvailStock) = txtAvailStock.Text
    End With
End Sub
Private Sub LoadMaterials()
    Dim strsql       As String

                                 
    strsql = "SELECT * From Materials Order By ItemDescription"
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    LVMaterials
              With lvwMaterials
                  If .ListItems.Count > 0 Then
                     Set .SelectedItem = .ListItems(1)
                     lvwMaterials_ItemClick .SelectedItem
                  End If
              End With
End Sub
Private Sub LVMaterials()
Dim MaterialsLI  As ListItem
lvwMaterials.ListItems.Clear
    With mmsADORst
        Do Until .EOF
            Set MaterialsLI = lvwMaterials.ListItems.Add(, , !ItemDescription & "")
            MaterialsLI.SubItems(M_Code) = !ItemCode & ""
            MaterialsLI.SubItems(M_Group) = !ItemGroup & ""
            MaterialsLI.SubItems(M_Unit) = !Unit & ""
            MaterialsLI.SubItems(M_Cost) = !Cost & ""
            MaterialsLI.SubItems(M_MinStock) = !MinStock & ""
            MaterialsLI.SubItems(M_AvailStock) = !AvailStock & ""
            MaterialsLI.SubItems(M_ItemID) = CStr(!ItemId)
            .MoveNext
        Loop
     End With
    Set MaterialsLI = Nothing
    Set mmsADORst = Nothing
End Sub
    
Private Sub ClearBox()
    CboGroup.Text = ""
    TxtMCode.Text = ""
    txtDes.Text = ""
    cboUnit.Text = ""
    txtCost.Text = ""
    txtMinStock.Text = ""
    txtAvailStock.Text = ""
    cboItemSearch.Text = ""
    txtItemSearch.Text = ""
End Sub
Private Sub BoxState(boxEnabled As Boolean)
    CboGroup.Enabled = boxEnabled
    txtDes.Enabled = boxEnabled
    cboUnit.Enabled = boxEnabled
    txtCost.Enabled = boxEnabled
    txtMinStock.Enabled = boxEnabled
    cboItemSearch.Enabled = Not boxEnabled
    txtItemSearch.Enabled = Not boxEnabled
End Sub
Private Sub ButtonState(buttonEnabled As Boolean)
    lvwMaterials.Enabled = buttonEnabled
    cmdAdd.Enabled = buttonEnabled
    cmdUpdate.Enabled = buttonEnabled
    cmdDelete.Enabled = buttonEnabled
    cmdExit.Enabled = buttonEnabled
    cmdSave.Enabled = Not buttonEnabled
    cmdCancel.Enabled = Not buttonEnabled
End Sub
Private Sub SendKeys_()
    SendKeys "{left}"
    SendKeys "{del}"
End Sub
Private Function DataValidation() As Boolean
  DataValidation = False
    If txtDes.Text = "" Then
         MsgBox "Fill-up Item Description.", vbExclamation, "Item Description Required"
         txtDes.SetFocus
        Exit Function
    End If
    If EncodeMode = "U" Then
    Else
        If DesExists Then
            MsgBox "Item Description Already Exist", vbExclamation, "Duplicate Description"
            txtDes.SetFocus
            Exit Function
        End If
     End If
    If CboGroup.Text = "" Then
        MsgBox "Fill-up Item Group.", vbExclamation, "Item Group Required"
        CboGroup.SetFocus
        Exit Function
    End If

    If cboUnit.Text = "" Then
        MsgBox "Fill-up Item Unit.", vbExclamation, "Item Unit Required"
        cboUnit.SetFocus
        Exit Function
    End If
    'If Val(txtCost.Text) = 0 Then
    '    MsgBox "Fill-up Item Cost.", vbExclamation, "Item Cost Required"
    '    txtCost.SetFocus
    '    Exit Function
    'End If
    If Val(txtMinStock.Text) = 0 Then
        MsgBox "Fill-up Item Minimum Stock.", vbExclamation, "Item Minimum Stock Required"
        txtMinStock.SetFocus
        Exit Function
    End If
    DataValidation = True

End Function
Private Function DesExists() As Boolean
    Dim objTempRst  As New ADODB.Recordset
    Dim strsql      As String

    strsql = "select count(*) as the_count from Materials where ItemDescription = '" & txtDes.Text & "'"
    objTempRst.Open strsql, mmsADOConn, adOpenForwardOnly, , adCmdText
    
    If objTempRst("the_count") > 0 Then
        DesExists = True
    Else
        DesExists = False
    End If

End Function
Private Function GetNextItemID() As Long
    mmsAdoCmd.CommandText = "SELECT MAX(ItemID) AS MaxID FROM Materials"
    Set mmsADORst = mmsAdoCmd.Execute
       
       If mmsADORst.EOF Then
           GetNextItemID = 1
          ElseIf IsNull(mmsADORst!MaxID) Then
               GetNextItemID = 1
       Else
           GetNextItemID = mmsADORst!MaxID + 1
       End If

    Set mmsADORst = Nothing

End Function
Public Function ButtonShortcuts()
On Error GoTo LocalError
   
   If ButtonPress = "A" Or ButtonPress = "a" Then
      CmdAdd_Click
    ElseIf ButtonPress = "U" Or ButtonPress = "u" Then
      CmdUpdate_Click
    ElseIf ButtonPress = "D" Or ButtonPress = "d" Then
      cmdDelete_Click
    ElseIf ButtonPress = "X" Or ButtonPress = "x" Then
      CmdExit_Click
    ElseIf ButtonPress = "S" Or ButtonPress = "s" Then
      cboItemSearch_GotFocus
      txtItemSearch.SetFocus
    Exit Function
    End If

LocalError:
    Exit Function
End Function


