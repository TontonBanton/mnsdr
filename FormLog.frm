VERSION 5.00
Begin VB.Form FormLog 
   Caption         =   "M&S COMPANY INCORPORATED"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   Icon            =   "FormLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4965
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtEncodeMode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1560
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   4665
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   2200
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2350
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   2200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtPWord 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   2625
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
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   2625
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
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   975
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
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "FormLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn              As ADODB.Connection
Private mmsAdoCmd               As ADODB.Command
Private mmsADORst               As ADODB.Recordset
Private strsql, EncodeMode      As String
Public UserName                 As String
Dim User, PWord                 As String
Private Sub Form_Load()
   ConnectToDB
End Sub
Private Sub txtUser_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
      txtPWord.SetFocus
   End If
End Sub
Private Sub txtUser_GotFocus()
  txtUser.Text = ""
End Sub
Private Sub txtPWord_KeyPress(KeyAscii As Integer)
   KeyAscii = ConvertUpper(KeyAscii)
   If KeyAscii = 13 Then
     cmdOk.SetFocus
   End If
End Sub
Private Sub txtPWord_GotFocus()
  txtPWord.Text = ""
End Sub
Private Sub cmdOk_Click()
    User = txtUser.Text
    PWord = txtPWord.Text
    strsql = "Select * from PWord Where UName like '" & User & "' and PWord like '" & PWord & "'"
    mmsAdoCmd.CommandText = strsql
    Set mmsADORst = mmsAdoCmd.Execute
    
    If mmsADORst.EOF = True Then
       MsgBox "INVALID ENTRY", vbCritical, "Security"
       txtUser.Text = ""
       txtPWord.Text = ""
       txtUser.SetFocus
    
        Else  '--- CORRECT PASSWORD
            If txtEncodeMode.Text = "SF" Then  ' -- CANCEL DR
                If MsgBox("CANCEL this DR_Number?", _
                    vbYesNo + vbQuestion, "Exit") = vbYes Then
                    'Unload Me
                    FormLog.Hide
                    FormDR.Enabled = True
                    FormDR.Show
                    Call FormDR.CancelDR
                Else
                    'Unload Me
                    FormLog.Hide
                    FormDR.Enabled = True
                    FormDR.Show
                End If
                ElseIf txtEncodeMode.Text = "PRODUCT" Then
                    'Unload Me
                    FormLog.Hide
                    FormProduct.txtEncodeMode = "PRODUCT"
                    FormProduct.Show
                ElseIf txtEncodeMode.Text = "PRODUCT2" Then
                    'Unload Me
                    FormLog.Hide
                    FormProduct.txtEncodeMode = "PRODUCT2"
                    FormProduct.Show
                ElseIf txtEncodeMode.Text = "SETTINGS" Then
                    'Unload Me
                    FormLog.Hide
                    FormSettings.Show
                
            Else '--- CORRECT PASSWORD "MainMenu"
               UserName = txtUser.Text
               'Unload Me
               FormLog.Hide
               FormMainMenu.Show
           End If
       txtUser.Text = ""
       txtPWord.Text = ""
    End If
   Set mmsADORst = Nothing
End Sub
Private Sub cmdCancel_Click()
If txtEncodeMode.Text = "SETTINGS" Or txtEncodeMode.Text = "PRODUCT" Then
    Unload Me
    FormMainMenu.Show
 ElseIf txtEncodeMode.Text = "SF" Then
    Unload Me
    FormDR.Enabled = True
    FormDR.Show
    FormDR.cmdSearch.Enabled = True
    FormDR.cmdSearch.SetFocus
 ElseIf txtEncodeMode.Text = "PRODUCT2" Then
    Unload Me
    FormDR.Enabled = True
    FormDR.Show
Else
    txtUser.Text = ""
    txtPWord.Text = ""
    txtUser.SetFocus
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
  cmdCancel_Click
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
Private Sub cmdOk_GotFocus()
   cmdOk.BackColor = &HC0FFC0
End Sub
Private Sub cmdOk_LostFocus()
   cmdOk.BackColor = &H8000000F
End Sub
Private Sub cmdCancel_GotFocus()
   cmdCancel.BackColor = &HC0FFC0
End Sub
Private Sub cmdCancel_LostFocus()
   cmdCancel.BackColor = &H8000000F
End Sub
