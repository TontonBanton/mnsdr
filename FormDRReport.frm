VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FormDRReport 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DELIVERY RECEIPT REPORTS"
   ClientHeight    =   10725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16905
   ControlBox      =   0   'False
   Icon            =   "FormDRReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10695
   ScaleMode       =   0  'User
   ScaleWidth      =   25171.86
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Height          =   1103
      Left            =   0
      TabIndex        =   13
      Top             =   903
      Width           =   5809
      Begin VB.TextBox txtYear 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         MaxLength       =   4
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cboMonth 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "FormDRReport.frx":1AC52
         Left            =   120
         List            =   "FormDRReport.frx":1AC7A
         TabIndex        =   1
         Text            =   "cboMonth"
         Top             =   350
         Width           =   2200
      End
      Begin MSMask.MaskEdBox txtDRStart 
         Height          =   420
         Left            =   2520
         TabIndex        =   2
         Top             =   350
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   741
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDREnding 
         Height          =   420
         Left            =   4100
         TabIndex        =   3
         Top             =   350
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   741
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000002&
      Height          =   1103
      Left            =   5820
      TabIndex        =   12
      Top             =   903
      Width           =   11055
      Begin VB.OptionButton OptDRActng 
         BackColor       =   &H80000002&
         Caption         =   "ACCOUNTING"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   360
         TabIndex        =   19
         Top             =   600
         Width           =   1785
      End
      Begin VB.OptionButton OptDRMarket 
         BackColor       =   &H80000002&
         Caption         =   "MARKET"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5640
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptDRArea 
         BackColor       =   &H80000002&
         Caption         =   "AREA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5640
         TabIndex        =   16
         Top             =   600
         Width           =   1100
      End
      Begin VB.OptionButton OptDRReport 
         BackColor       =   &H80000002&
         Caption         =   "REPORT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   330
         Width           =   1800
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   330
         Width           =   1800
      End
      Begin VB.OptionButton optDRProduct 
         BackColor       =   &H80000002&
         Caption         =   "PRODUCT"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3960
         TabIndex        =   7
         Top             =   600
         Width           =   1300
      End
      Begin VB.OptionButton OptDRDue 
         BackColor       =   &H80000002&
         Caption         =   "BUYER"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3960
         TabIndex        =   6
         Top             =   240
         Width           =   1100
      End
      Begin VB.OptionButton OptDRNumber 
         BackColor       =   &H80000002&
         Caption         =   "NUMBER"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2280
         TabIndex        =   5
         Top             =   600
         Width           =   1300
      End
      Begin VB.OptionButton OptDRSale 
         BackColor       =   &H80000002&
         Caption         =   "AR/SALE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   1300
      End
   End
   Begin MSComctlLib.ListView lvwDRSummary 
      Height          =   8715
      Left            =   0
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1920
      Width           =   16920
      _ExtentX        =   29845
      _ExtentY        =   15372
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   105
      Picture         =   "FormDRReport.frx":1ACE0
      Stretch         =   -1  'True
      Top             =   50
      Width           =   900
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
      Left            =   1080
      TabIndex        =   10
      Top             =   450
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
      Left            =   1080
      TabIndex        =   0
      Top             =   195
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   900
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   16935
   End
End
Attribute VB_Name = "FormDRReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mmsADOConn          As ADODB.Connection
Private mmsAdoCmd           As ADODB.Command
Private mmsADORst           As ADODB.Recordset
Private dbcommand           As ADODB.Command
Private strsql              As String
Dim DRSummaryLI             As ListItem

Dim ReportTransact, ReportType, ReportTittle, MonthSummary, StartMonth, EndMonth    As String
Dim DRGroup1, DRGroup1a, DRGroup2, DRGroup2a, DRGroup3, GetTotal                    As String
Dim QtyTotal, WtTotal, AmountTotal, CoShareTotal, IDNo                              As Currency
Private Sub Form_Load()
    Load Me
    ConnectToDB
      If (Month(Now)) = 1 Then
         cboMonth.Text = "DECEMBER"
      Else
         cboMonth.Text = UCase(MonthName(Month(Now) - 1))
      End If
   lblComp.Caption = FormMainMenu.lblComp.Caption
   lblHeader.Caption = FormMainMenu.lblHeader.Caption
   
End Sub
'---------------------------------------------------------------------------------
'                         C O N T R O L S   E V E N T S
'---------------------------------------------------------------------------------
Private Sub cboMonth_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtDRStart.SetFocus
   Else
      cboMonth.Text = ""
      SendKeys_
   End If
End Sub
Private Sub cboMonth_GotFocus()
      cboMonth.SelLength = Len(cboMonth.Text)
   If (Month(Now)) = 1 Then
       cboMonth.Text = "DECEMBER"
   Else
       cboMonth.Text = UCase(MonthName(Month(Now) - 1))
   End If
   lvwDRSummary.ListItems.Clear
   ClearOptions
End Sub
Private Sub cboMonth_Click()
   lvwDRSummary.ListItems.Clear
   ReportRange
End Sub
Private Sub cboMonth_LostFocus()
   If (Month(Now)) = 1 Then
       txtYear.Text = ((Year(Now) - 1))
   Else
       txtYear.Text = Year(Now)
   End If
   lvwDRSummary.ListItems.Clear
   ReportRange
End Sub

Private Sub txtDRStart_GotFocus()
 txtDRStart.SelLength = Len(txtDRStart.Text)
End Sub
Private Sub txtDRStart_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then
      Exit Sub
   ElseIf Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 Then
      'SendKeys_
   ElseIf KeyAscii = 13 Then
     If Not IsDate(txtDRStart.Text) Then
        MsgBox "Invalid Date.", vbExclamation, "Invalid Entry"
        txtDRStart.SetFocus
        txtDRStart.Text = Format$(StartMonth, "mm/dd/yyyy")
     Else
        txtDREnding.SetFocus
     End If
   End If
End Sub
Private Sub txtDREnding_GotFocus()
 txtDREnding.SelLength = Len(txtDREnding.Text)
End Sub
Private Sub txtDREnding_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then
      Exit Sub
   ElseIf Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 13 Then
      'SendKeys_
   ElseIf KeyAscii = 13 Then
     If Not IsDate(txtDREnding.Text) Then
        MsgBox "Invalid Date.", vbExclamation, "Invalid Entry"
        txtDREnding.SetFocus
        txtDREnding.Text = Format$(EndMonth, "mm/dd/yyyy")
     Else
        'txtYear.SetFocus
     End If
   End If
End Sub
Private Sub txtYear_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      OptDRSale.SetFocus
    Else
       If IsNumeric(Chr(KeyAscii)) Then
       Else
          SendKeys_
       End If
    End If
   lvwDRSummary.ListItems.Clear
End Sub
' ---------------------------------------------------------------------------
'                              B  U  T  T  O  N  S
' ---------------------------------------------------------------------------
Private Sub cmdPrint_Click()
    cmdExit.SetFocus
    ListViewPrint
End Sub
Private Sub cmdExit_Click()
   If MsgBox("Are you sure you want to exit?", _
        vbYesNo + vbQuestion, "Exit") = vbYes Then
     'ClearBox
     'ClearFrame
     Unload Me
     FormDR.Show
     FormDR.cmdExit.SetFocus
   Else
     Exit Sub
 End If
End Sub
' -------------------  D E L I V E R E D  -----------------------
'----------------------------------------------------------------
Private Sub OptDRSale_LostFocus()
   OptDRSale.FontBold = False
End Sub
Private Sub OptDRSale_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'Options
   End If
End Sub
Private Sub OptDRSale_GotFocus()

ReportType = "SALE"
StartMonth = txtDRStart.Text
EndMonth = txtDREnding.Text
OptDRSale.FontBold = True
'FrameHide
On Error GoTo LocalError
    
    SetlvwDRSale
    
    ReportType = "PRODUCT1"
    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "POMELO SUMMARY" & "")
    DRSummaryLI.ForeColor = &HC0&
    strsql = "SELECT Distinct DRProduct, DRCustomer From DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND DRProduct like '" & "POMELO" & "'"
    CommandExecute
    With mmsADORst
       Do Until .EOF
         Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
             DRSummaryLI.ForeColor = &HC0&
              DRSummaryLI.SubItems(1) = !DRCustomer
              DRGroup1 = !DRProduct
              DRGroup2 = !DRCustomer
              DRCustomerTotals
             .MoveNext
        Loop
    End With
    DRGroup1Totals
    '------------
    
    ReportType = "PRODUCT2"
    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
    DRSummaryLI.ForeColor = &HC0&
    
    strsql = "SELECT Distinct DRProduct, DRCustomer From DRDetails" _
           & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND DRProduct <> 'POMELO'"
    CommandExecute
    With mmsADORst
       Do Until .EOF
         Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , !DRProduct & "")
             DRSummaryLI.ForeColor = &HC0&
              DRSummaryLI.SubItems(1) = !DRCustomer
              DRGroup1 = !DRProduct
              DRGroup2 = !DRCustomer
              DRCustomerTotals
             .MoveNext
        Loop
    End With
    DRGroup1Totals
    
    '------------
    LVSpace
    ReportType = "SALE"
    strsql = "SELECT Distinct DRSaleTo From DRDetails" _
           & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# ORDER BY DRSaleTo"
    CommandExecute
    With mmsADORst
       Do Until .EOF
         Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , !DRSaleTo & "")
             DRSummaryLI.ForeColor = &HC0&
              DRGroup1 = !DRSaleTo
              DRReceivableList
              DRGroup1Totals
              LVSpace
             .MoveNext
        Loop
    End With
    DRSignatories
LocalError:
    Exit Sub
End Sub
Private Sub DRReceivableList()
    strsql = "SELECT Distinct DRCustomer, DRSaleTo From DRDetails" _
           & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
           & "  AND DRSaleTo like '" & DRGroup1 & "'"
    CommandExecute
    With mmsADORst
       Do Until .EOF
         Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
             DRSummaryLI.ForeColor = &HC0&
              DRSummaryLI.SubItems(1) = !DRCustomer
              DRGroup1 = !DRSaleTo
              DRGroup2 = !DRCustomer
              DRCustomerTotals
             .MoveNext
        Loop
        'DRReceivablesTotals
    End With
End Sub
Private Sub DRDeliveredSubItem()
                 strsql = "SELECT Distinct DRProduct FROM DRDetails" _
                          & "  WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
                          & "  AND DRSaleTo like '" & DRGroup1 & "'"
                  CommandExecute
                  lvwDRSummary.ForeColor = &H0&
                  With mmsADORst
                    Do Until .EOF
                    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
                        DRSummaryLI.SubItems(1) = !DRProduct
                        DRGroup2 = !DRProduct
                        DRGroup2Totals
                      .MoveNext

                    Loop
                 End With
End Sub
Private Sub DRCustomerTotals()
            If ReportType = "SALE" Then
                strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'" _
                        & "  AND DRCustomer like '" & DRGroup2 & "'"
                CommandExecute
                QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'" _
                        & "  AND DRCustomer like '" & DRGroup2 & "'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'" _
                        & "  AND DRCustomer like '" & DRGroup2 & "'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'" _
                        & "  AND DRCustomer like '" & DRGroup2 & "'"
                CommandExecute
                CoShareTotal = mmsADORst.Fields!Subtotal
            ElseIf ReportType = "PRODUCT1" Or ReportType = "PRODUCT2" Then
                        strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRProduct like '" & DRGroup1 & "'" _
                        & "  AND DRCustomer like '" & DRGroup2 & "'"
                CommandExecute
                QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRProduct like '" & DRGroup1 & "'" _
                        & "  AND DRCustomer like '" & DRGroup2 & "'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRProduct like '" & DRGroup1 & "'" _
                        & "  AND DRCustomer like '" & DRGroup2 & "'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRProduct like '" & DRGroup1 & "'" _
                        & "  AND DRCustomer like '" & DRGroup2 & "'"
                CommandExecute
                CoShareTotal = mmsADORst.Fields!Subtotal
            End If
                DRSummaryLI.SubItems(2) = Format$(QtyTotal, "#,###.#0")
                DRSummaryLI.SubItems(3) = Format$(WtTotal, "#,###.#0")
                DRSummaryLI.SubItems(4) = Format$(AmountTotal, "#,###.#0")
                DRSummaryLI.SubItems(5) = Format$(CoShareTotal, "#,###.#0")
End Sub
' -------------------  N U M B E R  -----------------------
'----------------------------------------------------------------
Private Sub OptDRNumber_LostFocus()
   OptDRNumber.FontBold = False
End Sub
Private Sub OptDRNumber_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'Options
   End If
End Sub
Private Sub OptDRNumber_GotFocus()

ReportType = "NUMBER"
StartMonth = txtDRStart.Text
EndMonth = txtDREnding.Text
OptDRNumber.FontBold = True
'FrameHide
On Error GoTo LocalError
    SetlvwDRSale
    strsql = "SELECT Distinct DRNum, DRDate, DRCustomer, DRNumDetails, DRSaleTo From DRDetails" _
           & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
           & " ORDER BY DRNum"
    CommandExecute
    With mmsADORst
       Do Until .EOF
         Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , !DRNum & " - " & !DRDate & " ( " & !DRCustomer & " )")
              DRSummaryLI.SubItems(2) = !DRNumDetails
              DRSummaryLI.ForeColor = &HC0&
              DRGroup1 = !DRNum
              DRGroup2 = !DRSaleTo
              DRNumberSubItem
              LVSpace
             .MoveNext
        Loop
    End With
     Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
         DRSummaryLI.SubItems(1) = "                          TOTAL ---->"
     DRGroup1 = "B3 CASH SALES"
     DRNumberTotals
     DRGroup1 = "B3 ACCOUNT RECEIVABLES"
     DRNumberTotals
     DRGroup1 = "B3 REPRESENTATION"
     DRNumberTotals
     DRSignatories
LocalError:
    Exit Sub
End Sub
Private Sub DRNumberSubItem()
                 strsql = "SELECT DRProductClass, DRSackBox, DRQty, DRWeight, DRCost, DRAmount FROM DRDetails" _
                          & "  WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
                          & "  AND DRNum like '" & DRGroup1 & "'"
                  CommandExecute
                  lvwDRSummary.ForeColor = &H0&
                  With mmsADORst
                    Do Until .EOF
                    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
                    If DRGroup2 = "B3 CASH SALES" Then
                        DRSummaryLI.SubItems(1) = !DRProductClass
                        DRSummaryLI.SubItems(3) = !DRSackBox
                        DRSummaryLI.SubItems(4) = !DRQty
                        DRSummaryLI.SubItems(5) = !DRWeight
                        DRSummaryLI.SubItems(6) = !DRCost
                        DRSummaryLI.SubItems(7) = !DRAmount
                    End If
                    If DRGroup2 = "B3 ACCOUNT RECEIVABLES" Then
                        DRSummaryLI.SubItems(1) = !DRProductClass
                        DRSummaryLI.SubItems(3) = !DRSackBox
                        DRSummaryLI.SubItems(8) = !DRQty
                        DRSummaryLI.SubItems(9) = !DRWeight
                        DRSummaryLI.SubItems(10) = !DRCost
                        DRSummaryLI.SubItems(11) = !DRAmount
                    End If
                    If DRGroup2 = "B3 REPRESENTATION" Then
                        DRSummaryLI.SubItems(1) = !DRProductClass
                        DRSummaryLI.SubItems(3) = !DRSackBox
                        DRSummaryLI.SubItems(12) = !DRWeight
                    End If
                        'DRGroup2 = !DRProductClass
                        'DRGroup2Totals
                      .MoveNext
                    Loop
                 End With
End Sub
Private Sub DRNumberTotals()
                strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'"
                CommandExecute
                QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'"
                CommandExecute
                CoShareTotal = mmsADORst.Fields!Subtotal
                
               
               If DRGroup1 = "B3 CASH SALES" Then
                 DRSummaryLI.SubItems(4) = Format$(QtyTotal, "#,###.#0")
                 DRSummaryLI.SubItems(5) = Format$(WtTotal, "#,###.#0")
                 DRSummaryLI.SubItems(7) = Format$(AmountTotal, "#,###.#0")
               End If
               If DRGroup1 = "B3 ACCOUNT RECEIVABLES" Then
                 DRSummaryLI.SubItems(8) = Format$(QtyTotal, "#,###.#0")
                 DRSummaryLI.SubItems(9) = Format$(WtTotal, "#,###.#0")
                 DRSummaryLI.SubItems(11) = Format$(AmountTotal, "#,###.#0")
               End If
               If DRGroup1 = "B3 REPRESENTATION" Then
                 DRSummaryLI.SubItems(12) = Format$(WtTotal, "#,###.#0")
               End If
End Sub
' -------------------  D E S T I N A T I O N  -------------------------------
'----------------------------------------------------------------------------
Private Sub SetlvwDRDue()
    lvwDRSummary.ListItems.Clear
    With lvwDRSummary
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , " ", .Width * 0.05
        .ColumnHeaders.Add , , "BUYER ", .Width * 0.15
        .ColumnHeaders.Add , , "NUMBER ", .Width * 0.08
        .ColumnHeaders.Add , , "CLASSIFICATION ", .Width * 0.15
        .ColumnHeaders.Add , , "BLOCK ", .Width * 0.08
        .ColumnHeaders.Add , , "QTY", .Width * 0.08
        .ColumnHeaders.Add , , "WEIGHT ", .Width * 0.08
        .ColumnHeaders.Add , , "AMOUNT", .Width * 0.08
    End With
         lvwDRSummary.ColumnHeaders.Item(6).Alignment = lvwColumnRight
         lvwDRSummary.ColumnHeaders.Item(7).Alignment = lvwColumnRight
         lvwDRSummary.ColumnHeaders.Item(8).Alignment = lvwColumnRight
End Sub
Private Sub OptDRDue_LostFocus()
   OptDRDue.FontBold = False
End Sub
Private Sub OptDRDue_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'Options
   End If
End Sub
Private Sub OptDRDue_GotFocus()
ReportType = "DESTINATION"
StartMonth = txtDRStart.Text
EndMonth = txtDREnding.Text
OptDRDue.FontBold = True
'FrameHide
On Error GoTo LocalError
    SetlvwDRDue
    
strsql = "SELECT Distinct DRCustomer From DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# ORDER BY DRCustomer"
CommandExecute
    With mmsADORst
       Do Until .EOF
         Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , !DRCustomer & "")
              DRSummaryLI.ForeColor = &HC0&
              DRGroup1 = !DRCustomer
              DRCustomerSubItem
              DRDueTotals
              LVSpace
             .MoveNext
        Loop
    End With
LocalError:
    Exit Sub
End Sub
Private Sub DRCustomerSubItem()
    strsql = "SELECT DRCustomer, DRNum, DRClass, DRAreaDetails, DRQty, DRWeight, DRAmount, DRProduct FROM DRDetails" _
           & "  WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
           & "   AND DRCustomer like '" & DRGroup1 & "'" _
           & "  ORDER BY DRProduct, DRClass "
           
                  CommandExecute
                  lvwDRSummary.ForeColor = &H0&
                  With mmsADORst
                    Do Until .EOF
                    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , !DRProduct & "")
                        DRSummaryLI.SubItems(1) = !DRCustomer
                        DRSummaryLI.SubItems(2) = !DRNum
                        DRSummaryLI.SubItems(3) = !DRClass
                        DRSummaryLI.SubItems(4) = !DRAreaDetails
                        DRSummaryLI.SubItems(5) = !DRQty
                        DRSummaryLI.SubItems(6) = !DRWeight
                        DRSummaryLI.SubItems(7) = !DRAmount
                      .MoveNext
                    Loop
                 End With
End Sub
Private Sub DRDueTotals()
                strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & " AND DRCustomer like '" & DRGroup1 & "'"
                CommandExecute
                QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & " AND DRCustomer like '" & DRGroup1 & "'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & " AND DRCustomer like '" & DRGroup1 & "'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & " AND DRCustomer like '" & DRGroup1 & "'"
                CommandExecute
                CoShareTotal = mmsADORst.Fields!Subtotal
                
                    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
                        DRSummaryLI.SubItems(4) = "                       TOTAL ---->"
                        DRSummaryLI.SubItems(5) = Format$(QtyTotal, "#,###.#0")
                        DRSummaryLI.SubItems(6) = Format$(WtTotal, "#,###.#0")
                        DRSummaryLI.SubItems(7) = Format$(AmountTotal, "#,###.#0")
                        DRSummaryLI.ListSubItems(4).Bold = True
                        DRSummaryLI.ListSubItems(4).ForeColor = &HC0&
                        DRSummaryLI.ListSubItems(5).Bold = True
                        DRSummaryLI.ListSubItems(5).ForeColor = &HC0&
                        DRSummaryLI.ListSubItems(6).Bold = True
                        DRSummaryLI.ListSubItems(6).ForeColor = &HC0&
                        DRSummaryLI.ListSubItems(7).Bold = True
                        DRSummaryLI.ListSubItems(7).ForeColor = &HC0&
                
End Sub
' -------------------   P R O D U C T   -----------------------
'----------------------------------------------------------------
Private Sub OptDRProduct_LostFocus()
   optDRProduct.FontBold = False
End Sub
Private Sub OptDRProduct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'Options
   End If
End Sub
Private Sub OptDRProduct_GotFocus()

ReportType = "PRODUCT"
StartMonth = txtDRStart.Text
EndMonth = txtDREnding.Text
optDRProduct.FontBold = True
'FrameHide
On Error GoTo LocalError
SetlvwDRSale
strsql = "SELECT Distinct DRProduct From DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# ORDER BY DRProduct"
CommandExecute
    With mmsADORst
       Do Until .EOF
         Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , !DRProduct & "")
              DRSummaryLI.ForeColor = &HC0&
              DRGroup1 = !DRProduct
              DRProductSubItem
              DRGroup1Totals
              LVSpace
             .MoveNext
        Loop
    End With
LocalError:
    Exit Sub
End Sub
Private Sub DRProductSubItem()
strsql = "SELECT Distinct DRClass FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND DRProduct like '" & DRGroup1 & "'"
CommandExecute
lvwDRSummary.ForeColor = &H0&
With mmsADORst
        Do Until .EOF
        Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
        DRSummaryLI.SubItems(1) = !DRClass
        DRGroup2 = !DRClass
        DRGroup2Totals
        .MoveNext
        Loop
End With
End Sub
' -------------------   M  A  R  K  E T   -----------------------
'---------------------------_-------------------------------------
Private Sub OptDRMarket_LostFocus()
   OptDRMarket.FontBold = False
End Sub
Private Sub OptDRMarket_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'Options
   End If
End Sub
Private Sub OptDRMarket_GotFocus()
StartMonth = txtDRStart.Text
EndMonth = txtDREnding.Text
OptDRMarket.FontBold = True
'FrameHide
On Error GoTo LocalError
    
    SetlvwDRSale
    
    '-------------------- M A R K E T --------
    ReportType = "MARKET"
    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "SUMMARY BY CUSTOMER" & "")
    DRSummaryLI.ForeColor = &HC0&
    
    strsql = "SELECT Distinct DRDestination From DRDetails" _
           & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
           & " ORDER BY DRDestination"
    CommandExecute
    With mmsADORst
       Do Until .EOF
              DRGroup1 = !DRDestination
              DRDestinationSubItem
              DRGroup1Totals
              LVSpace
             .MoveNext
        Loop
    End With
    
     '-------------------- M A R K E T --------
    ReportType = "MARKET2"
    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "SUMMARY BY CLASSIFICATION" & "")
    DRSummaryLI.ForeColor = &HC0&
    
    strsql = "SELECT Distinct DRDestination From DRDetails" _
           & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
           & " ORDER BY DRDestination"
    CommandExecute
    With mmsADORst
       Do Until .EOF
              DRGroup1 = !DRDestination
              DRDestinationSubItem
              DRGroup1Totals
              LVSpace
             .MoveNext
        Loop
    End With
    
   

LocalError:
    Exit Sub
End Sub
Private Sub DRDestinationSubItem()
        If ReportType = "MARKET" Then
                  strsql = "SELECT Distinct DRCustomer FROM DRDetails" _
                          & "  WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
                          & "  AND DRDestination like '" & DRGroup1 & "'"
                  CommandExecute
                  lvwDRSummary.ForeColor = &H0&
                  With mmsADORst
                    Do Until .EOF
                    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , DRGroup1 & "")
                        DRSummaryLI.SubItems(1) = !DRCustomer
                        DRGroup2 = !DRCustomer
                        DRGroup2Totals
                      .MoveNext
                    Loop
                  End With
        End If
        
         If ReportType = "MARKET2" Then
                  strsql = "SELECT Distinct DRProduct, DRClass FROM DRDetails" _
                          & "  WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
                          & "  AND DRDestination like '" & DRGroup1 & "'"
                  CommandExecute
                  lvwDRSummary.ForeColor = &H0&
                  With mmsADORst
                    Do Until .EOF
                    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , DRGroup1 & "")
                        DRSummaryLI.SubItems(1) = !DRProduct & " - " & !DRClass
                        DRGroup2 = !DRProduct
                        DRGroup3 = !DRClass
                        DRGroup2Totals
                      .MoveNext
                    Loop
                  End With
        End If
               
End Sub
' -------------------   A    R    E    A   -----------------------
'---------------------------_-------------------------------------
Private Sub OptDRArea_LostFocus()
   OptDRArea.FontBold = False
End Sub
Private Sub OptDRArea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'Options
   End If
End Sub
Private Sub OptDRArea_GotFocus()
StartMonth = txtDRStart.Text
EndMonth = txtDREnding.Text
OptDRArea.FontBold = True
'FrameHide
On Error GoTo LocalError
    
    SetlvwDRSale
    
    ReportType = "AREA3"
    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "DETAILS BY PHASE/BLOCK CLASSIFICATION" & "")
    DRSummaryLI.ForeColor = &HC0&
    
    strsql = "SELECT Distinct DRBlock From DRDetails" _
           & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
           & " ORDER BY DRBlock"
    CommandExecute
    With mmsADORst
       Do Until .EOF
              DRGroup1 = !DRBlock
              DRAreaSubItem
              DRGroup1Totals
              LVSpace
             .MoveNext
        Loop
    End With
    
    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "" & "")
    DRSummaryLI.SubItems(2) = "---------- END OF PHASE/BLOCK SUMMARY ----------"
    

LocalError:
    Exit Sub
End Sub
Private Sub DRAreaSubItem()
        If ReportType = "AREA3" Then
                   strsql = "SELECT Distinct DRClass FROM DRDetails" _
                          & "  WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
                          & "  AND DRBlock like '" & DRGroup1 & "'"
                  CommandExecute
                  lvwDRSummary.ForeColor = &H0&
                  With mmsADORst
                    Do Until .EOF
                    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , DRGroup1 & "")
                        DRSummaryLI.SubItems(1) = !DRClass
                        DRGroup2 = !DRClass
                        DRGroup2Totals
                      .MoveNext
                    Loop
                  End With
        End If
        
End Sub
Private Sub DRGroup1Totals()
            If ReportType = "SALE" Then         '----------- SALE
                strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'"
                CommandExecute
                QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'"
                CommandExecute
                CoShareTotal = mmsADORst.Fields!Subtotal
          End If
          If ReportType = "PRODUCT" Or ReportType = "AREA" Or ReportType = "AREA2" Then
            strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                   & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# AND DRProduct like '" & DRGroup1 & "'"
                   CommandExecute
                   QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                   & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# AND DRProduct like '" & DRGroup1 & "'"
                   CommandExecute
                   CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                   & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# AND DRProduct like '" & DRGroup1 & "'"
                   CommandExecute
                  AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                   & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# AND DRProduct like '" & DRGroup1 & "'"
                   CommandExecute
                   CoShareTotal = mmsADORst.Fields!Subtotal
          End If
          
          If ReportType = "PRODUCT1" Then
            strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails " _
                   & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND DRProduct like '" & "POMELO" & "'"
                   CommandExecute
                   QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                   & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND DRProduct like '" & "POMELO" & "'"
                   CommandExecute
                   CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                   & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND DRProduct like '" & "POMELO" & "'"
                   CommandExecute
                  AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                   & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND DRProduct like '" & "POMELO" & "'"
                   CommandExecute
                   CoShareTotal = mmsADORst.Fields!Subtotal
          End If
          
          If ReportType = "PRODUCT2" Then
               strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails " _
                   & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND DRProduct <> 'POMELO'"
                   CommandExecute
                   QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                   & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND DRProduct <> 'POMELO'"
                   CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                   & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND DRProduct <> 'POMELO'"
                   CommandExecute
                  AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                   & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND DRProduct <> 'POMELO'"
                   CommandExecute
                   CoShareTotal = mmsADORst.Fields!Subtotal
          End If
                     
          If ReportType = "AREA3" Then
                strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRBlock like '" & DRGroup1 & "'"
                CommandExecute
                QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRBlock like '" & DRGroup1 & "'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                       & "  AND DRBlock like '" & DRGroup1 & "'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRBlock like '" & DRGroup1 & "'"
                CommandExecute
                CoShareTotal = mmsADORst.Fields!Subtotal
          End If
          
          If ReportType = "MARKET" Or ReportType = "MARKET2" Then
                strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRDestination like '" & DRGroup1 & "'"
                CommandExecute
                QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRDestination like '" & DRGroup1 & "'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRDestination like '" & DRGroup1 & "'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRDestination like '" & DRGroup1 & "'"
                CommandExecute
                CoShareTotal = mmsADORst.Fields!Subtotal
          End If
          
          If ReportType = "AREA4" Then
                strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRClass like '" & DRGroup1 & "'"
                CommandExecute
                QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRClass like '" & DRGroup1 & "'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRClass like '" & DRGroup1 & "'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRClass like '" & DRGroup1 & "'"
                CommandExecute
                CoShareTotal = mmsADORst.Fields!Subtotal
          End If
                FinalTotals

End Sub
Private Sub DRGroup2Totals()
    If ReportType = "SALE" Then
        strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                & "  AND DRSaleTo like '" & DRGroup1 & "' AND DRProduct like '" & DRGroup2 & "'"
         CommandExecute
         QtyTotal = mmsADORst.Fields!Subtotal
         strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
         & "  AND DRSaleTo like '" & DRGroup1 & "' AND DRProduct like '" & DRGroup2 & "'"
         CommandExecute
         WtTotal = mmsADORst.Fields!Subtotal
         strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                & "  AND DRSaleTo like '" & DRGroup1 & "' AND DRProduct like '" & DRGroup2 & "'"
        CommandExecute
        AmountTotal = mmsADORst.Fields!Subtotal
        strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
               & "  AND DRSaleTo like '" & DRGroup1 & "' AND DRProduct like '" & DRGroup2 & "'"
        CommandExecute
        CoShareTotal = mmsADORst.Fields!Subtotal
    End If
    
    If ReportType = "PRODUCT" Then
         strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                & "  AND DRProduct like '" & DRGroup1 & "' AND DRClass like '" & DRGroup2 & "'"
         CommandExecute
         QtyTotal = mmsADORst.Fields!Subtotal
         strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                & "  AND DRProduct like '" & DRGroup1 & "' AND DRClass like '" & DRGroup2 & "'"
         CommandExecute
         WtTotal = mmsADORst.Fields!Subtotal
         strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                & "  AND DRProduct like '" & DRGroup1 & "' AND DRClass like '" & DRGroup2 & "'"
        CommandExecute
        AmountTotal = mmsADORst.Fields!Subtotal
        strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                & "  AND DRProduct like '" & DRGroup1 & "' AND DRClass like '" & DRGroup2 & "'"
        CommandExecute
        CoShareTotal = mmsADORst.Fields!Subtotal
    End If
            If ReportType = "PRODUCT2" Then
                strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRProduct like '" & DRGroup1 & "'"
                CommandExecute
                QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRProduct like '" & DRGroup1 & "'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRProduct like '" & DRGroup1 & "'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRProduct like '" & DRGroup1 & "'"
                CommandExecute
                CoShareTotal = mmsADORst.Fields!Subtotal
            End If
            If ReportType = "AREA" Then
                strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRProduct like '" & DRGroup1 & "'" _
                        & "  AND DRPhase like '" & DRGroup2 & "'"
                CommandExecute
                QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRProduct like '" & DRGroup1 & "'" _
                        & "  AND DRPhase like '" & DRGroup2 & "'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRProduct like '" & DRGroup1 & "'" _
                        & "  AND DRPhase like '" & DRGroup2 & "'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRProduct like '" & DRGroup1 & "'" _
                        & "  AND DRPhase like '" & DRGroup2 & "'"
                CommandExecute
                CoShareTotal = mmsADORst.Fields!Subtotal
          End If
          If ReportType = "AREA3" Then
                strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRBlock like '" & DRGroup1 & "'" _
                        & "  AND DRClass like '" & DRGroup2 & "'"
                CommandExecute
                QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRBlock like '" & DRGroup1 & "'" _
                        & "  AND DRClass like '" & DRGroup2 & "'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRBlock like '" & DRGroup1 & "'" _
                        & "  AND DRClass like '" & DRGroup2 & "'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRBlock like '" & DRGroup1 & "'" _
                        & "  AND DRClass like '" & DRGroup2 & "'"
                CommandExecute
                CoShareTotal = mmsADORst.Fields!Subtotal
          End If
          
          If ReportType = "AREA4" Then
                strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRClass like '" & DRGroup1 & "'" _
                        & "  AND DRPhase like '" & DRGroup2 & "'"
                CommandExecute
                QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRClass like '" & DRGroup1 & "'" _
                        & "  AND DRPhase like '" & DRGroup2 & "'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRClass like '" & DRGroup1 & "'" _
                        & "  AND DRPhase like '" & DRGroup2 & "'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRClass like '" & DRGroup1 & "'" _
                        & "  AND DRPhase like '" & DRGroup2 & "'"
                CommandExecute
                CoShareTotal = mmsADORst.Fields!Subtotal
          End If
          
          If ReportType = "MARKET" Then
                strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRDestination like '" & DRGroup1 & "'" _
                        & "  AND DRCustomer like'" & DRGroup2 & "'"
                CommandExecute
                QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRDestination like '" & DRGroup1 & "'" _
                        & "  AND DRCustomer like'" & DRGroup2 & "'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRDestination like '" & DRGroup1 & "'" _
                        & "  AND DRCustomer like'" & DRGroup2 & "'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRDestination like '" & DRGroup1 & "'" _
                        & "  AND DRCustomer like'" & DRGroup2 & "'"
                CommandExecute
                CoShareTotal = mmsADORst.Fields!Subtotal
          End If
           If ReportType = "MARKET2" Then
                strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRDestination like '" & DRGroup1 & "'" _
                        & "  AND DRProduct like'" & DRGroup2 & "'" _
                        & "  AND DRClass like'" & DRGroup3 & "'"
                CommandExecute
                QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRDestination like '" & DRGroup1 & "'" _
                        & "  AND DRProduct like'" & DRGroup2 & "'" _
                        & "  AND DRClass like'" & DRGroup3 & "'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRDestination like '" & DRGroup1 & "'" _
                        & "  AND DRProduct like'" & DRGroup2 & "'" _
                        & "  AND DRClass like'" & DRGroup3 & "'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRDestination like '" & DRGroup1 & "'" _
                        & "  AND DRProduct like'" & DRGroup2 & "'" _
                        & "  AND DRClass like'" & DRGroup3 & "'"
                CommandExecute
                CoShareTotal = mmsADORst.Fields!Subtotal
          End If
                DRSummaryLI.SubItems(2) = Format$(QtyTotal, "#,###.#0")
                DRSummaryLI.SubItems(3) = Format$(WtTotal, "#,###.#0")
                DRSummaryLI.SubItems(4) = Format$(AmountTotal, "#,###.#0")
                DRSummaryLI.SubItems(5) = Format$(CoShareTotal, "#,###.#0")
End Sub
'---------------------------------------------------------
Private Sub OptDRActng_LostFocus()
   OptDRActng.FontBold = False
End Sub
Private Sub OptDRActng_GotFocus()
ReportType = "ACCOUNTING"
StartMonth = txtDRStart.Text
EndMonth = txtDREnding.Text
OptDRActng.FontBold = True
'FrameHide
On Error GoTo LocalError
    
    SetlvwDRActng
    
    strsql = "SELECT * From DRDetails" _
           & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
           & " ORDER BY DRNum"
    CommandExecute
    With mmsADORst
       Do Until .EOF
       Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , !DRID & "")
             DRSummaryLI.SubItems(1) = !DRDate
             DRSummaryLI.SubItems(2) = !DRCustomer
             DRSummaryLI.SubItems(3) = !DRProduct & " - " & !DRClass
             DRSummaryLI.SubItems(4) = !DRQty
             DRSummaryLI.SubItems(5) = !DRWeight
             DRSummaryLI.SubItems(6) = !DRCost
             DRSummaryLI.SubItems(7) = !DRAmount
             DRSummaryLI.SubItems(8) = !DRCoShare
             DRSummaryLI.SubItems(9) = !DRSaleTo
             .MoveNext
        Loop
    End With
    AccountingTotals
     'DRSignatories
LocalError:
    Exit Sub
End Sub
Private Sub AccountingTotals()
    strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# "
    CommandExecute
    QtyTotal = mmsADORst.Fields!Subtotal
    strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# "
    CommandExecute
    WtTotal = mmsADORst.Fields!Subtotal
    strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# "
    CommandExecute
    AmountTotal = mmsADORst.Fields!Subtotal
    strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# "
    CommandExecute
    CoShareTotal = mmsADORst.Fields!Subtotal

    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
        DRSummaryLI.SubItems(4) = Format$(QtyTotal, "#,###.#0")
        DRSummaryLI.SubItems(5) = Format$(WtTotal, "#,###.#0")
        'DRSummaryLI.SubItems(6) = Format$(AmountTotal, "#,###.#0")
        DRSummaryLI.SubItems(7) = Format$(AmountTotal, "#,###.#0")
        DRSummaryLI.SubItems(8) = Format$(CoShareTotal, "#,###.#0")
        DRSummaryLI.ListSubItems(4).Bold = True: DRSummaryLI.ListSubItems(4).ForeColor = &HC0&
        DRSummaryLI.ListSubItems(5).Bold = True: DRSummaryLI.ListSubItems(5).ForeColor = &HC0&
        'DRSummaryLI.ListSubItems(6).Bold = True: DRSummaryLI.ListSubItems(6).ForeColor = &HC0&
        DRSummaryLI.ListSubItems(7).Bold = True: DRSummaryLI.ListSubItems(7).ForeColor = &HC0&
        DRSummaryLI.ListSubItems(8).Bold = True: DRSummaryLI.ListSubItems(8).ForeColor = &HC0&

End Sub
' ---------------------------------------------------------------------------------------
'                                         R E P O R T
'----------------------------------------------------------------------------------------
Private Sub OptDRReport_LostFocus()
   OptDRReport.FontBold = False
End Sub
Private Sub OptDRReport_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'Options
   End If
End Sub
Private Sub OptDRReport_GotFocus()

ReportType = "REPORT"
StartMonth = txtDRStart.Text
EndMonth = txtDREnding.Text
OptDRReport.FontBold = True
'FrameHide
On Error GoTo LocalError
    
    SetlvwDRReport
        
   
    'SUNGEE
    strsql = "SELECT Distinct DRCustomer,DRRef From DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
           & " AND DRCustomer like 'SUNGEE' ORDER By DRCustomer "
    CommandExecute
    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "SUNGEE ACCOUNT")
    With mmsADORst
       Do Until .EOF
              DRGroup1 = !DRCustomer
              DRGroup2 = !DRRef
              DRReportSubItem1
             .MoveNext
        Loop
    End With
    strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND DRCustomer like 'SUNGEE'"
    CommandExecute
    WtTotal = mmsADORst.Fields!Subtotal
    strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND DRCustomer like 'SUNGEE'"
    CommandExecute
    AmountTotal = mmsADORst.Fields!Subtotal
    AccountTotal
    
    'SM ACCOUNT
    strsql = "SELECT Distinct DRCustomer,DRRef FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
          & " AND DRCustomer like 'SM%'ORDER By DRCustomer "
    CommandExecute
    LVSpace
    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "SM ACCOUNT")
    With mmsADORst
       Do Until .EOF
              DRGroup1 = !DRCustomer
              DRGroup2 = !DRRef
              DRReportSubItem1
             .MoveNext
        Loop
    End With
        strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND DRCustomer like 'SM%'"
        CommandExecute
        WtTotal = mmsADORst.Fields!Subtotal
        strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "# AND DRCustomer like 'SM%'"
        CommandExecute
        AmountTotal = mmsADORst.Fields!Subtotal
        AccountTotal
                
    'DCO ACCOUNT
    strsql = "SELECT Distinct DRProduct,DRCustomer,DRRef From DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
           & " AND DRProduct like 'POMELO' AND DRCustomer NOT like 'SM%' AND DRCustomer NOT like 'SUNGEE'AND DRCustomer NOT like 'M&S MAL%' ORDER By DRCustomer "
    CommandExecute
    LVSpace
    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "DCO ACCOUNT")
    With mmsADORst
       Do Until .EOF
              DRGroup1 = !DRCustomer
              DRGroup2 = !DRRef
              DRReportSubItem1
             .MoveNext
      Loop
    End With
    strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
           & " AND DRProduct NOT like 'MANGO%' AND DRCustomer NOT like 'SM%' AND DRCustomer NOT like 'SUNGEE' AND DRCustomer NOT like 'M&S MAL%'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                       & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
                       & " AND DRProduct NOT like 'MANGO%'" _
                       & " AND DRCustomer NOT like 'SM%'" _
                       & " AND DRCustomer NOT like 'SUNGEE'" _
                       & " AND DRCustomer NOT like 'M&S MAL%'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                AccountTotal

    'COOP ACCOUNT
    strsql = "SELECT Distinct DRCustomer,DRRef FROM DRDetails" _
           & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
           & " AND DRCustomer like 'M&S MAL%'" _
           & " ORDER By DRCustomer "
    CommandExecute
    LVSpace
    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "COOP ACCOUNT")
    With mmsADORst
       Do Until .EOF
              DRGroup1 = !DRCustomer
              DRGroup2 = !DRRef
              DRReportSubItem1
             .MoveNext
       Loop
    End With
                
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                       & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
                       & " AND DRCustomer like 'M&S MAL%'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                       & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
                       & " AND DRCustomer like 'M&S MAL%'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                AccountTotal

    'MANGO/MISC
    strsql = "SELECT DISTINCT DRProduct,DRCustomer,DRRef FROM DRDetails" _
           & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
           & " AND DRProduct NOT like 'POM%'" _
           & " ORDER By DRCustomer "
    CommandExecute
    LVSpace
    With mmsADORst
    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "" & "")
       Do Until .EOF
              DRGroup1 = !DRCustomer
              DRGroup2 = !DRRef
              DRReportSubItem1
             .MoveNext
       Loop
    End With
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                       & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
                       & " AND DRProduct NOT like 'POM%'"
                CommandExecute
                WtTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                       & " WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
                       & " AND DRProduct NOT like 'POM%'"
                CommandExecute
                AmountTotal = mmsADORst.Fields!Subtotal
                AccountTotal
     DRSignatories
LocalError:
    Exit Sub
End Sub
Private Sub DRReportSubItem1()
Dim rownum, i As Integer

        strsql = "SELECT * FROM DRDetails" _
                       & "  WHERE DRDate BETWEEN #" & StartMonth & "# And #" & EndMonth & "#" _
                       & "  AND DRCustomer like '" & DRGroup1 & "'" _
                       & "  AND DRREf like '" & DRGroup2 & "'" _
                       & "  ORDER BY DRRef"
        CommandExecute
        lvwDRSummary.ForeColor = &H0&
                  
        With mmsADORst
                         'GET NUMBER OF ROWS
                          rownum = 0
                          Do While Not .EOF
                            rownum = rownum + 1
                           .MoveNext
                          Loop
            
            If rownum = 1 Then  'SINGLE ROW
                   .MoveFirst
                       Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , !DRProduct & "")
                         DRSummaryLI.SubItems(1) = !DRPhase
                         DRSummaryLI.SubItems(2) = !DRSONum
                         DRSummaryLI.SubItems(3) = !DRDate
                         DRSummaryLI.SubItems(4) = !DRRef
                         DRSummaryLI.SubItems(5) = !DRCustomer
                         DRSummaryLI.SubItems(6) = !DRClass
                         DRSummaryLI.SubItems(7) = !DRWeight
                         DRSummaryLI.SubItems(8) = !DRCost
                         DRSummaryLI.SubItems(9) = !DRAmount
                         DRSummaryLI.SubItems(10) = !DRARNum
                         DRSummaryLI.SubItems(11) = !DRCINum
                         DRSummaryLI.SubItems(12) = !DRCHINum
                         DRSummaryLI.SubItems(14) = !DRTotalWeight
                         DRSummaryLI.SubItems(15) = !DRTotalAmount
                         
                         
                         
                         '.MoveNext
            Else
                   .MoveFirst
                       Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , !DRProduct & "")
                         DRSummaryLI.SubItems(1) = !DRPhase
                         DRSummaryLI.SubItems(2) = !DRSONum
                         DRSummaryLI.SubItems(3) = !DRDate
                         DRSummaryLI.SubItems(4) = !DRRef
                         DRSummaryLI.SubItems(5) = !DRCustomer
                         DRSummaryLI.SubItems(6) = !DRClass
                         DRSummaryLI.SubItems(7) = !DRWeight
                         DRSummaryLI.SubItems(8) = !DRCost
                         DRSummaryLI.SubItems(9) = !DRAmount
                         .MoveNext
                  
                   rownum = rownum - 1
                   i = 1
                     Do While i < rownum
                        Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "" & "")
                         DRSummaryLI.SubItems(1) = !DRPhase
                         DRSummaryLI.SubItems(6) = !DRClass
                         DRSummaryLI.SubItems(7) = !DRWeight
                         DRSummaryLI.SubItems(8) = !DRCost
                         DRSummaryLI.SubItems(9) = !DRAmount
                         .MoveNext
                         i = i + 1
                     Loop
                      'LAST ROW
                       Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "" & "")
                         DRSummaryLI.SubItems(6) = !DRClass
                         DRSummaryLI.SubItems(7) = !DRWeight
                         DRSummaryLI.SubItems(8) = !DRCost
                         DRSummaryLI.SubItems(9) = !DRAmount
                         DRSummaryLI.SubItems(10) = !DRARNum
                         DRSummaryLI.SubItems(11) = !DRCINum
                         DRSummaryLI.SubItems(12) = !DRCHINum
                         DRSummaryLI.SubItems(14) = !DRTotalWeight
                         DRSummaryLI.SubItems(15) = !DRTotalAmount
            End If
        End With
                 
End Sub
Private Sub AccountTotal()
                         Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
                         DRSummaryLI.SubItems(13) = "TOTAL > > > > > "
                         DRSummaryLI.SubItems(14) = Format$(WtTotal, "#,###.#0")
                         DRSummaryLI.SubItems(15) = Format$(AmountTotal, "#,###.#0")
End Sub

'---------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------

Private Sub DRReceivablesTotals()
                strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'"
                mmsAdoCmd.CommandText = strsql
                Set mmsADORst = mmsAdoCmd.Execute
                QtyTotal = mmsADORst.Fields!Subtotal
                strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'"
                mmsAdoCmd.CommandText = strsql
                Set mmsADORst = mmsAdoCmd.Execute
                WtTotal = mmsADORst.Fields!Subtotal
                Set mmsADORst = Nothing
                strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'"
                mmsAdoCmd.CommandText = strsql
                Set mmsADORst = mmsAdoCmd.Execute
                AmountTotal = mmsADORst.Fields!Subtotal
                Set mmsADORst = Nothing
                strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails  " _
                        & " WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# " _
                        & "  AND DRSaleTo like '" & DRGroup1 & "'"
                mmsAdoCmd.CommandText = strsql
                Set mmsADORst = mmsAdoCmd.Execute
                CoShareTotal = mmsADORst.Fields!Subtotal
                Set mmsADORst = Nothing
                LVSpace
                FinalTotals
End Sub
Private Sub DRTotal()
    strsql = " SELECT SUM(DRQty) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# "
    CommandExecute
    QtyTotal = mmsADORst.Fields!Subtotal
    strsql = " SELECT SUM(DRWeight) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# "
    CommandExecute
    WtTotal = mmsADORst.Fields!Subtotal
    Set mmsADORst = Nothing
    strsql = " SELECT SUM(DRAmount) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# "
    CommandExecute
    AmountTotal = mmsADORst.Fields!Subtotal
    Set mmsADORst = Nothing
    strsql = " SELECT SUM(DRCoShare) as SubTotal FROM DRDetails WHERE DRDate BETWEEN #" & StartMonth & "# and #" & EndMonth & "# "
    CommandExecute
    CoShareTotal = mmsADORst.Fields!Subtotal
    Set mmsADORst = Nothing
    FinalTotals
End Sub
Private Sub FinalTotals()
                    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
                        DRSummaryLI.SubItems(1) = "                       TOTAL ---->"
                        DRSummaryLI.SubItems(2) = Format$(QtyTotal, "#,###.#0")
                        DRSummaryLI.SubItems(3) = Format$(WtTotal, "#,###.#0")
                        DRSummaryLI.SubItems(4) = Format$(AmountTotal, "#,###.#0")
                        DRSummaryLI.SubItems(5) = Format$(CoShareTotal, "#,###.#0")
                        DRSummaryLI.ListSubItems(1).Bold = True
                        DRSummaryLI.ListSubItems(1).ForeColor = &HC0&
                        DRSummaryLI.ListSubItems(2).Bold = True
                        DRSummaryLI.ListSubItems(2).ForeColor = &HC0&
                        DRSummaryLI.ListSubItems(3).Bold = True
                        DRSummaryLI.ListSubItems(3).ForeColor = &HC0&
                        DRSummaryLI.ListSubItems(4).Bold = True
                        DRSummaryLI.ListSubItems(4).ForeColor = &HC0&
                        DRSummaryLI.ListSubItems(5).Bold = True
                        DRSummaryLI.ListSubItems(5).ForeColor = &HC0&
End Sub
Private Sub DRSignatories()
                    LVSpace
                    LVSpace
                    LVSpace
                    If ReportType = "SALE" Then
                    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
                        DRSummaryLI.SubItems(1) = "PREPARED BY"
                        DRSummaryLI.SubItems(2) = "AUDITED BY"
                        DRSummaryLI.SubItems(3) = "NOTED BY"
                        DRSummaryLI.SubItems(5) = "APPROVED BY"
                    LVSpace
                    LVSpace
                    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
                        DRSummaryLI.SubItems(1) = FormSettings.txtPrepared.Text
                        DRSummaryLI.SubItems(2) = FormSettings.txtAudited.Text
                        DRSummaryLI.SubItems(3) = FormSettings.txtNoted.Text
                        DRSummaryLI.SubItems(5) = FormSettings.txtApproved.Text
                   End If
                   If ReportType = "NUMBER" Then
                    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
                        DRSummaryLI.SubItems(1) = "PREPARED BY"
                        DRSummaryLI.SubItems(4) = "AUDITED BY"
                        DRSummaryLI.SubItems(8) = "NOTED BY"
                        DRSummaryLI.SubItems(11) = "APPROVED BY"
                    LVSpace
                    LVSpace
                    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , "")
                        DRSummaryLI.SubItems(1) = FormSettings.txtPrepared.Text
                        DRSummaryLI.SubItems(4) = FormSettings.txtAudited.Text
                        DRSummaryLI.SubItems(8) = FormSettings.txtNoted.Text
                        DRSummaryLI.SubItems(11) = FormSettings.txtApproved.Text
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
Private Sub CommandExecute()
        mmsAdoCmd.CommandText = strsql
        Set mmsADORst = mmsAdoCmd.Execute
End Sub
Private Sub ReportRange()
Dim LastDay As Long
On Error GoTo LocalError
   MonthSummary = Month(CDate("1 " & cboMonth.Text))
   StartMonth = MonthSummary & "/" & "1" & "/" & txtYear.Text
   LastDay = Day(DateSerial(Year(txtYear.Text), Month(StartMonth) + 1, 0))
   EndMonth = MonthSummary & "/" & LastDay & "/" & txtYear.Text
   txtDRStart.Text = Format$(StartMonth, "mm/dd/yyyy")
   txtDREnding.Text = Format$(EndMonth, "mm/dd/yyyy")
LocalError:
    Exit Sub
End Sub
Private Sub SetlvwDRSale()
    lvwDRSummary.ListItems.Clear
    
    If ReportType = "NUMBER" Then
        With lvwDRSummary
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "DR#-DATE", .Width * 0.2
        .ColumnHeaders.Add , , "BUYER", .Width * 0.15
        .ColumnHeaders.Add , , "OR#", .Width * 0.05
        .ColumnHeaders.Add , , "SACKS", .Width * 0.05
        .ColumnHeaders.Add , , "QTY", .Width * 0.05
        .ColumnHeaders.Add , , "WT", .Width * 0.05
        .ColumnHeaders.Add , , "U/P", .Width * 0.05
        .ColumnHeaders.Add , , "AMOUNT", .Width * 0.05
        .ColumnHeaders.Add , , "QTY", .Width * 0.05
        .ColumnHeaders.Add , , "WT", .Width * 0.05
        .ColumnHeaders.Add , , "U/P", .Width * 0.05
        .ColumnHeaders.Add , , "AMOUNT", .Width * 0.05
        .ColumnHeaders.Add , , "REP. KGS.", .Width * 0.05
    End With
     lvwDRSummary.ColumnHeaders.Item(4).Alignment = lvwColumnRight
     lvwDRSummary.ColumnHeaders.Item(5).Alignment = lvwColumnRight
     lvwDRSummary.ColumnHeaders.Item(6).Alignment = lvwColumnRight
     lvwDRSummary.ColumnHeaders.Item(7).Alignment = lvwColumnRight
     lvwDRSummary.ColumnHeaders.Item(8).Alignment = lvwColumnRight
   
   Else
   
    With lvwDRSummary
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , " ", .Width * 0.1
        .ColumnHeaders.Add , , " ", .Width * 0.15
        .ColumnHeaders.Add , , "QTY", .Width * 0.15
        .ColumnHeaders.Add , , "WEIGHT ", .Width * 0.15
        .ColumnHeaders.Add , , "AMOUNT", .Width * 0.15
        .ColumnHeaders.Add , , "CO.SHARE", .Width * 0.15
        .ColumnHeaders.Add , , " ", .Width * 0#
    End With
         lvwDRSummary.ColumnHeaders.Item(3).Alignment = lvwColumnRight
         lvwDRSummary.ColumnHeaders.Item(4).Alignment = lvwColumnRight
         lvwDRSummary.ColumnHeaders.Item(5).Alignment = lvwColumnRight
         lvwDRSummary.ColumnHeaders.Item(6).Alignment = lvwColumnRight
  End If
End Sub
Private Sub SetlvwDRActng()
    lvwDRSummary.ListItems.Clear
        With lvwDRSummary
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "NUMBER", .Width * 0.08
        .ColumnHeaders.Add , , " ", .Width * 0.08
        .ColumnHeaders.Add , , "BUYER", .Width * 0.15
        .ColumnHeaders.Add , , "PRODUCT", .Width * 0.2
        .ColumnHeaders.Add , , "QTY", .Width * 0.08
        .ColumnHeaders.Add , , "WT", .Width * 0.08
        .ColumnHeaders.Add , , "PRICE", .Width * 0.08
        .ColumnHeaders.Add , , "AMOUNT", .Width * 0.08
        .ColumnHeaders.Add , , "COSHARE", .Width * 0.08
        .ColumnHeaders.Add , , "REMARK", .Width * 0.2
    End With
     lvwDRSummary.ColumnHeaders.Item(5).Alignment = lvwColumnRight
     lvwDRSummary.ColumnHeaders.Item(6).Alignment = lvwColumnRight
     lvwDRSummary.ColumnHeaders.Item(7).Alignment = lvwColumnRight
     lvwDRSummary.ColumnHeaders.Item(8).Alignment = lvwColumnRight
     lvwDRSummary.ColumnHeaders.Item(9).Alignment = lvwColumnRight
End Sub
Private Sub SetlvwDRReport()
    lvwDRSummary.ListItems.Clear
        With lvwDRSummary
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "PRODUCT", .Width * 0.05
        .ColumnHeaders.Add , , "AREA", .Width * 0.05
        .ColumnHeaders.Add , , "SO#", .Width * 0.05
        .ColumnHeaders.Add , , "DATE", .Width * 0.05
        .ColumnHeaders.Add , , "DR#", .Width * 0.05
        .ColumnHeaders.Add , , "BUYER", .Width * 0.12
        .ColumnHeaders.Add , , "CLASS", .Width * 0.13
        .ColumnHeaders.Add , , "WT", .Width * 0.05
        .ColumnHeaders.Add , , "U/P", .Width * 0.05
        .ColumnHeaders.Add , , "AMOUNT", .Width * 0.05
        .ColumnHeaders.Add , , "AR#", .Width * 0.04
        .ColumnHeaders.Add , , "CI#", .Width * 0.04
        .ColumnHeaders.Add , , "CH.I#", .Width * 0.04
        .ColumnHeaders.Add , , "CR#", .Width * 0.04
        .ColumnHeaders.Add , , "WT", .Width * 0.07
        .ColumnHeaders.Add , , "AMOUNT", .Width * 0.07
    End With
     lvwDRSummary.ColumnHeaders.Item(8).Alignment = lvwColumnRight
     lvwDRSummary.ColumnHeaders.Item(9).Alignment = lvwColumnRight
     lvwDRSummary.ColumnHeaders.Item(10).Alignment = lvwColumnRight
     lvwDRSummary.ColumnHeaders.Item(15).Alignment = lvwColumnRight
     lvwDRSummary.ColumnHeaders.Item(16).Alignment = lvwColumnRight
End Sub
Private Sub ListViewPrint()
    Dim ExcelObj   As Object
    Dim ExcelBook  As Object
    Dim ExcelSheet As Object
    Dim lst As ListItem, lst1 As ListSubItem, row As Integer, col As Integer, i As Integer
    Dim AppExcel   As Variant

    Set AppExcel = CreateObject("Excel.application")
    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelBook = ExcelObj.WorkBooks.Add
    Set ExcelSheet = ExcelBook.Worksheets(1)
    
    With ExcelObj.activesheet
          .Pagesetup.Orientation = 1 'LANDSCAPE xlPortrait 2
          '.Pagesetup.RightHeader = "Date: " & "&D" & " " & Format(Time, "hh:mm") & " Page: " & "&P of &N"
          .Pagesetup.LeftMargin = 15
          .Pagesetup.RightMargin = 15
          .Pagesetup.TopMargin = 20
          .Pagesetup.BottomMargin = 20
    End With
    
    With ExcelSheet
          .Pagesetup.RightFooter = " Page: " & "&P"
          .range("A:Z").Font.Size = 10
          .range("A:Z").RowHeight = 13
          .range("A:A").Font.Bold = True
          .range("A1").Value = lblComp.Caption
          .range("A2").Value = FormMainMenu.lblHeader
          .range("A4").Value = "DELIVERY RECEIPT " & ReportType & " REPORT"
          .range("A5").Value = "REPORT PERIOD  :  " & StartMonth & "-" & EndMonth
          
         If ReportType = "SALE" Or ReportType = "DESTINATION" Or ReportType = "PRODUCT" Or ReportType = "AREA3" Or ReportType = "MARKET2" Then
            .range("C:G").NumberFormat = "#,##0.00"
            .Columns.columnwidth = 15
            .Columns(1).columnwidth = 10
            .Columns(2).columnwidth = 25
            .Columns(3).columnwidth = 10
         End If
                    
                    '------------ NUMBER
                    If ReportType = "NUMBER" Then
                       With ExcelObj.activesheet
                            .Pagesetup.Orientation = 2 'LANDSCAPE xlPortrait 2
                       End With
                    .range("C:Z").NumberFormat = "#,##0.00"
                    .Columns.columnwidth = 8
                    .Columns(1).columnwidth = 3
                    .Columns(2).columnwidth = 35
                    .Columns(8).columnwidth = 10
                    .Columns(12).columnwidth = 10
                    End If
                       
                    '------------ REPORT
                    If ReportType = "REPORT" Then
                       With ExcelObj.activesheet
                            .Pagesetup.Orientation = 2 'LANDSCAPE xlPortrait 2
                       End With
                    .range("H:P").NumberFormat = "#,##0.00"
                    .Columns.columnwidth = 7
                    .Columns(4).columnwidth = 10
                    .Columns(10).columnwidth = 10
                    .Columns(6).columnwidth = 20
                    .Columns(7).columnwidth = 20
                    .Columns(15).columnwidth = 12
                    .Columns(16).columnwidth = 12
                    End If
                    
                    '------------ ACCOUNTING
                    If ReportType = "ACCOUNTING" Then
                       With ExcelObj.activesheet
                            .Pagesetup.Orientation = 1 'LANDSCAPE xlPortrait 2
                       End With
                    .range("A:Z").Font.Size = 9
                    .range("E:H").NumberFormat = "#,##0.00"
                    .Columns.columnwidth = 8
                    .Columns(1).columnwidth = 5
                    .Columns(3).columnwidth = 15
                    .Columns(4).columnwidth = 18
                    .Columns(10).columnwidth = 18
                    End If
    End With
    
    row = 7
    col = 1
    
     For i = 1 To lvwDRSummary.ColumnHeaders.Count
        ExcelSheet.cells(row, col) = lvwDRSummary.ColumnHeaders(i)
        col = col + 1
      Next
       row = 9
       col = 1
        For Each lst In lvwDRSummary.ListItems
          col = 1
          ExcelSheet.cells(row, col) = lst.Text
          col = col + 1
          For Each lst1 In lst.ListSubItems
            ExcelSheet.cells(row, col) = lst1.Text
            col = col + 1
          Next
          row = row + 1
        Next

    ExcelObj.Visible = True
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing
    
End Sub
Private Sub SetNothing()
    Set DRSummaryLI = Nothing
    Set mmsADORst = Nothing
End Sub
Private Sub ClearOptions()
   OptDRSale.Value = False
   OptDRNumber.Value = False
End Sub
Private Sub LVSpace()
    Set DRSummaryLI = lvwDRSummary.ListItems.Add(, , " ")
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
