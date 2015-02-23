Public Class clsTest
    Inherits Quasi97.clsQSTTestNET

    Public Shared SharedTestID$ = "Pause"
    Public Property Message$ = "Click OK when ready to continue"
    Private dbPTR As OleDb.OleDbConnection
    Private myTableName$ = "PauseTestModule_Settings"

    Public Overrides Function CheckRecords(NewDBase As String) As System.Collections.Generic.List(Of Short)
        Return MyBase.GenericCheckRecords(NewDBase, myTableName)
    End Function

    Public Overrides Sub ClearResults(Optional doRefreshPlot As Boolean = False)
        'do nothing
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
            Return 0
        End Get
    End Property

    Public Overrides Sub RemoveRecord()
        MyBase.GenericRemoveRecord(dbPTR, myTableName)
        dbPTR = Nothing
    End Sub

    Public Overrides Sub RestoreParameters()
        If dbPTR Is Nothing Then Return
        Try
            dbPTR.Open()
            Dim sqlCom As New OleDb.OleDbCommand("SELECT * FROM " & myTableName & " WHERE Setup=" & Setup, dbPTR)
            Dim rsParams As OleDb.OleDbDataReader = sqlCom.ExecuteReader
            If rsParams.Read Then
                Message = rsParams("Message")
            End If

        Catch ex As Exception
            MsgBox("Restoreparams " & ex.Message)
        Finally
            If dbPTR.State <> ConnectionState.Closed Then dbPTR.Close()
        End Try
    End Sub

    Public Overrides Sub RunTest()
        If QST.QSTHardware.DualChanMode Then Return
        If MsgBox(Message, vbOKCancel, SharedTestID) = MsgBoxResult.Cancel Then
            QST.QuasiParameters.AbortTest = True
        End If
    End Sub

    Public Overrides Sub SetDBase(ByRef NewDBase As String, Optional ByRef voidParam As Object = Nothing)
        MyBase.GenericSetDBase(dbPTR, NewDBase)
    End Sub

    Public Overrides Sub StoreParameters()
        If dbPTR Is Nothing Then Return
        Try
            dbPTR.Open()
            Dim sqlCom As New OleDb.OleDbCommand("", dbPTR)
            sqlCom.CommandText = "DELETE * FROM " & myTableName & " WHERE Setup=" & Setup
            sqlCom.ExecuteNonQuery()
            sqlCom.CommandText = "INSERT INTO " & myTableName & " (Setup, Message) VALUES (" & Setup & ",'" & Message & "')"
            sqlCom.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("Storeparams " & ex.Message)
        Finally
            If dbPTR.State <> ConnectionState.Closed Then dbPTR.Close()
        End Try
    End Sub

    Public Overrides ReadOnly Property TestID As String
        Get
            Return sharedtestid
        End Get
    End Property

    Public Sub New()
        MyBase.New()
        colParameters.Add(New Quasi97.clsTestParam("Message", "Message", Me, Message.GetType))
    End Sub
End Class
