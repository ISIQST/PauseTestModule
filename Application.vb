Public Class Application
    Implements IDisposable

    Private assyName$ = "PauseTestModule"
    Public Property ModuleDescr() As String = "pause test module"
    Public Property ModuleID() As String = assyName
    Public TemplateDB$ = "PauseTestModule.mda"
    Public Manufacturer$ = "Integral Solutions Int'l"
    Public AppDataPath$ = ""

    Private Sub GetAppDataPath()
        Dim strAppData As String
        strAppData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) 'GetSpecialFolderA(mceIDLPaths.CSIDL_COMMON_APPDATA)
        strAppData = strAppData & "\" & Manufacturer
        AppDataPath = strAppData & "\Quasi97"
    End Sub

    Public Sub Initialize2(ByRef q As Object)
        QST = q
        GetAppDataPath()
        'add tables needed for pause test module
        Quasi97.ADOUtils.SynchronizeDatabases(AppDataPath & "\" & TemplateDB, QST.QuasiParameters.SetupTmpFileName)
        QST.QuasiParameters.RegisterTestClassNET(clsTest.SharedTestID, assyName, assyName & ".clsTest", My.Resources.Hourglass, "Quasi97.ucGenericNoGraph")

    End Sub

    Public ReadOnly Property CustomStressSupport As Boolean
        Get
            Return False
        End Get
    End Property

    Sub Dispose() Implements System.IDisposable.Dispose
        If Not QST Is Nothing Then
            Call QST.QuasiParameters.UnregisterTestClass(clsTest.SharedTestID)
        End If

        QST = Nothing
    End Sub

    Public ReadOnly Property QuasiAddIn() As Boolean
        Get
            QuasiAddIn = True
        End Get
    End Property


End Class
