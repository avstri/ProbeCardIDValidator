Imports System.Collections.Generic
Public Class clsTest
    Inherits Quasi97.clsQSTTestNET

    Public Shared SharedTestID As String = "PCard ID Check"
    Private myTable As String = "ProbeCardIDValidator_Params"
    Const myResultName$ = "Result"

    Private dbPTR As OleDb.OleDbConnection
    Public Property CompatibleIDs$ = "" 'semicolon separated
    Public Property AbortOnFail As Boolean = True

    Public Overrides Function CheckRecords(NewDBase As String) As System.Collections.Generic.List(Of Short)
        If NewDBase <> "" Then
            'check if table exists, and if not create it
            Dim ldb As OleDb.OleDbConnection = Nothing
            MyBase.GenericSetDBase(ldb, NewDBase)
            Dim sq As New OleDb.OleDbCommand("SELECT * FROM " & myTable & " WHERE SETUP=" & Setup.ToString, ldb)
            Try
                ldb.Open()
                'check if table exists, if not, create it
                Try
                    sq.ExecuteNonQuery()
                Catch ex As Exception
                    'exception means there is no table
                    sq.CommandText = "CREATE TABLE [" & myTable & "] ([Setup] smallint,[CompatibleIDs] Text(255),[AbortOnFail] yesno)"
                    Try
                        sq.ExecuteNonQuery()
                    Catch eex As Exception
                        MsgBox(eex.Message)
                    End Try
                End Try
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                If ldb.State <> ConnectionState.Closed Then ldb.Close()
            End Try
        End If
        Return MyBase.GenericCheckRecords(NewDBase, myTable)
    End Function

    Public Overrides Sub ClearResults(Optional doRefreshPlot As Boolean = False)
        colResults.Clear()
    End Sub

    Public Overrides ReadOnly Property ContainsGraph As Short
        Get
            Return 0
        End Get
    End Property

    Public Overrides ReadOnly Property ContainsResultPerCycle As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overrides ReadOnly Property DualChannelCapable As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides ReadOnly Property FeatureVector As UInteger
        Get
            Return Features.AdaptiveParameters Or Features.parametersregistered Or Features.localAbort 'Or Features.SingleThread
        End Get
    End Property

    Public Overrides Sub RemoveRecord()
        MyBase.GenericRemoveRecord(dbPTR, mytable)
    End Sub

    Public Overrides Sub RestoreParameters()
        If dbPTR Is Nothing Then Return
        Try
            dbPTR.Open()
            Dim sql As New OleDb.OleDbCommand("SELECT * FROM " & myTable & " WHERE Setup=" & Setup.ToString, dbPTR)
            Dim rs As OleDb.OleDbDataReader = sql.ExecuteReader
            If Not rs.Read() Then
                rs.Close()
                Dim sqwr As New OleDb.OleDbCommand("INSERT INTO " & myTable & " (Setup,CompatibleIDs,AbortOnFail) VALUES (" & Setup.ToString & ",'',yes)", dbPTR)
                sqwr.ExecuteNonQuery()
                rs = sql.ExecuteReader
                rs.Read()
            End If

            CompatibleIDS = rs("CompatibleIDs").ToString
            AbortOnFail = CBool(rs("AbortOnFail"))

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            dbPTR.Close()
        End Try
    End Sub

    Public Overrides Sub RunTest()
        Try
            Dim rslt As New Quasi97.ResultNet
            ClearResults(False)
            colResults.Add(rslt)
            If ProbeCardValid Then
                rslt.AddResult(myResultName, "1", 1, False)
            Else
                rslt.AddResult(myResultName, "0", 1, False)
                If AbortOnFail Then Application.QST.QuasiParameters.AbortTest = True
            End If
            rslt.AddInfo(Me, 0, Application.QST.QuasiParameters.CurInfo)
            MyBase.RaiseNewInfoAvailable()
            MyBase.RaiseNewResultsAvailable({0})

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Overrides Sub SetDBase(ByRef NewDBase As String, Optional ByRef voidParam As Object = Nothing)
        MyBase.GenericSetDBase(dbPTR, NewDBase)
        If (NewDBase <> "") And (Setup <> 0) Then
            'the following is optional
            AddHandler Application.QST.QuasiParameters.StartStopDeviceInitiate, AddressOf mystartStop
        Else
            RemoveHandler Application.QST.QuasiParameters.StartStopDeviceInitiate, AddressOf mystartStop
        End If
    End Sub

    Public Overrides Sub StoreParameters()
        If dbPTR Is Nothing Then Return
        Try
            dbPTR.Open()
            Dim sql As New OleDb.OleDbCommand("DELETE * FROM " & myTable & " WHERE Setup=" & Setup.ToString, dbPTR)
            sql.ExecuteNonQuery()
            Dim sqwr As New OleDb.OleDbCommand("INSERT INTO " & myTable & " (Setup,CompatibleIDs,abortonfail) VALUES (" & Setup.ToString & ",'" & CompatibleIDs & "'," & AbortOnFail & ")", dbPTR)
            sqwr.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            dbPTR.Close()
        End Try
    End Sub

    Public Overrides ReadOnly Property TestID As String
        Get
            Return sharedtestid
        End Get
    End Property

    Public Sub New()
        MyBase.New()
        colParameters.Add(New Quasi97.clsTestParam("Compatible IDs", "CompatibleIDs", Me, CompatibleIDS.GetType, False, False))
        colParameters.Add(New Quasi97.clsTestParam("Abort On Fail", "AbortOnFail", Me, AbortOnFail.GetType, False, False))
        colResultNames.Add(New Quasi97.ResultName(myResultName)) '1 for pass, 0 for fail

    End Sub

    Sub mystartStop(ByRef doStart As Boolean)
        If doStart Then
            If Not ProbeCardValid() Then
                Call Application.QST.UserMsgBox("This setup was created for a different probe card " & vbNewLine & _
                                                 Application.QST.Barx2G3Parameters.PREEprom1.EEGetField(Quasi97.cls2xBarG3.EEPROBECARDFLDS.BRDPARTID).ToUpper & _
                                                 " is not among [" & CompatibleIDs & "]", vbOKOnly, SharedTestID, 0)
                'doStart = False 'optional
            End If
        End If
    End Sub

    Protected Overrides Sub Finalize()
        RemoveHandler Application.QST.QuasiParameters.StartStopDeviceInitiate, AddressOf mystartStop
        MyBase.Finalize()
    End Sub

    Private Function ProbeCardValid() As Boolean
        Dim parsedOut() As String = CompatibleIDs.ToUpper.Split(";")
        If parsedOut.Length = 0 Then
            Return True
        Else
            Dim curPCardID$ = Application.QST.Barx2G3Parameters.PREEprom1.EEGetField(Quasi97.cls2xBarG3.EEPROBECARDFLDS.BRDPARTID).ToUpper.Trim
            If parsedOut.Contains(curPCardID) Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

End Class
