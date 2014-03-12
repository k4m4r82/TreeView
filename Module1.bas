Attribute VB_Name = "Module1"
Option Explicit

Public conn    As ADODB.Connection
Public strSql  As String

Public Function KonekToServer() As Boolean
    Dim strConn As String
    
    On Error GoTo errHandle
            
    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\database.mdb"
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = strConn
    conn.Open
    
    KonekToServer = True
    
    Exit Function

errHandle:
    KonekToServer = False
End Function

Public Function openRecordset(ByVal query As String) As ADODB.Recordset
    Dim obj As ADODB.Recordset
    
    Set obj = New ADODB.Recordset
    obj.CursorLocation = adUseClient
    obj.Open query, conn, adOpenForwardOnly, adLockReadOnly
    Set openRecordset = obj
End Function

Public Sub closeRecordset(ByVal vRs As ADODB.Recordset)
    If Not (vRs Is Nothing) Then
        If vRs.State = adStateOpen Then vRs.Close
    End If
    
    Set vRs = Nothing
End Sub

Public Function getRecordCount(ByVal vRs As ADODB.Recordset) As Long
    vRs.MoveLast
    getRecordCount = vRs.RecordCount
    vRs.MoveFirst
End Function

