Public Class Application
    Implements Quasi97.iAddin

    Friend Shared QST As Quasi97.Application
    Public Property ModuleID As String = "ProbeCardIDValidator" Implements Quasi97.iAddin.ModuleID

    Public ReadOnly Property CustomStressSupport As Boolean Implements Quasi97.iAddin.CustomStressSupport
        Get
            Return False
        End Get
    End Property

    Public ReadOnly Property QuasiAddIn As String Implements Quasi97.iAddin.QuasiAddIn
        Get
            Return True
        End Get
    End Property

    Public Sub Initialize2(ByRef q As Quasi97.Application) Implements Quasi97.iAddin.Initialize2
        QST = q
        QST.QuasiParameters.RegisterTestClassNET(clsTest.sharedtestid, ModuleID, ModuleID & ".clsTest", Nothing, "Quasi97.ucGenericNoGraph")
    End Sub

    Public Sub RunCustomStress(sM As Quasi97.clsStress.StressMode, UniqueEventID As Short, ByRef EventParams As String) Implements Quasi97.iAddin.RunCustomStress
        'do nothing
    End Sub

    Public Sub ValidateStressParam(eID As Short, ByRef ParamValue As String, ByRef RetValue As Boolean) Implements Quasi97.iAddin.ValidateStressParam
        'do nothing
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If qst IsNot Nothing Then
                    QST.QuasiParameters.UnregisterTestClass(clsTest.sharedtestid)
                    QST = Nothing
                End If
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
