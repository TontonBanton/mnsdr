Attribute VB_Name = "Module1"
Public Function ConvertUpper(pintKeyValue As Integer) As Integer
'  Common function to force alphabetic keyboard characters to uppercase
'  when called from the KeyPress event.
'  Typical call:
'      KeyAscii = ConvertUpper(KeyAscii)
    If Chr$(pintKeyValue) >= "a" And Chr$(pintKeyValue) <= "z" Then
        pintKeyValue = pintKeyValue - 32
    End If
    ConvertUpper = pintKeyValue
End Function
'-----------------------------------------------------------------------------Private Sub ConnectToDB()
Public Function ConnectToDB()
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

End Function

