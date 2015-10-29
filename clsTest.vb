Public Class clsTest
    Inherits Quasi97.clsSimpleTest

    Public Shared SharedTestID$ = "Pause"
    Public Property Message$ = "Click OK when ready to continue"
    Private myTableName$ = "PauseTestModule_Settings"

    Public Overrides Sub RunTest()
        If QST.QSTHardware.DualChanMode Then Return
        If MsgBox(Message, vbOKCancel, SharedTestID) = MsgBoxResult.Cancel Then
            QST.QuasiParameters.AbortTest = True
        End If
    End Sub

    Public Overrides ReadOnly Property TestID As String
        Get
            Return SharedTestID
        End Get
    End Property

    Public Sub New()
        MyBase.New()
        colParameters.Add(New Quasi97.clsTestParam("Message", "Message", Me, Message.GetType))
    End Sub

    Public Overrides ReadOnly Property MyTable As String
        Get
            Return myTableName
        End Get
    End Property
End Class
