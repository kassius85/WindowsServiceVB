Public Class ProjectInstaller

    Public Sub New()
        MyBase.New()

        'El Diseñador de componentes requiere esta llamada.
        InitializeComponent()

        'Agregue el código de inicialización después de llamar a InitializeComponent

    End Sub

    Protected Overrides Sub OnAfterInstall(savedState As IDictionary)
        MyBase.OnAfterInstall(savedState)

        'The following code starts the services after it is installed.
        Using serviceController As New ServiceProcess.ServiceController(IntervalService.ServiceName)
            serviceController.Start()
        End Using
    End Sub
End Class