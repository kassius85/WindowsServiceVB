Imports System.IO
Imports System.Threading
Imports System.Configuration

Public Class IntervalTaskService
    Private GuardaLog As Boolean = (ConfigurationManager.AppSettings("MuestraLog") = "S")
    'Get the Scheduled Time from AppSettings.
    Private ScheduledTime As Date = Date.Parse(ConfigurationManager.AppSettings("ScheduledTime"))

    Private Schedular As Timer

    'INICIO DEL TIMER
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Agregue el código aquí para iniciar el servicio. Este método debería poner
        ' en movimiento los elementos para que el servicio pueda funcionar.
        ScheduleService()
    End Sub

    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
        Schedular.Dispose()
    End Sub

    Public Sub ScheduleService()
        Try
            'Initialize the Schedular
            Schedular = New Timer(New TimerCallback(AddressOf SchedularCallback))

            'Get the interval mode
            Dim mode As String = ConfigurationManager.AppSettings("Mode").ToUpper()

            Dim schedule As String = String.Empty
            Dim executeTask As Boolean = True

            'Get the current minute
            Dim tempDate As Date = Date.Now
            Dim currentDate As Date = New Date(tempDate.Year, tempDate.Month, tempDate.Day, tempDate.Hour, tempDate.Minute, 0, 0)

            Select Case mode
                Case "DAILY"
                    'If Scheduled Time is passed set Schedule for the next day.
                    If currentDate <> ScheduledTime Then
                        If currentDate > ScheduledTime Then ScheduledTime = ScheduledTime.AddDays(1)

                        executeTask = False
                    Else
                        ScheduledTime = ScheduledTime.AddDays(1)
                    End If

                Case "INTERVAL"
                    'Get the Interval in Minutes from AppSettings.
                    Dim intervalMinutes As Integer = Convert.ToInt32(ConfigurationManager.AppSettings("IntervalMinutes"))

                    If intervalMinutes > 0 Then
                        If currentDate <> ScheduledTime Then
                            If currentDate > ScheduledTime Then ScheduledTime = ScheduledTime.AddDays(1)

                            executeTask = False
                        Else
                            'Set the Scheduled Time by adding the Interval to Current Time.
                            ScheduledTime = currentDate.AddMinutes(intervalMinutes)

                            'If Scheduled Time is passed set Schedule for the next Interval.
                            If Date.Now > ScheduledTime Then ScheduledTime = ScheduledTime.AddMinutes(intervalMinutes)
                        End If
                    Else
                        Throw New Exception("El intervalo en minutos debe ser mayor a cero.")
                    End If

                Case Else
                    Throw New Exception("El modo definido no es válido.")

            End Select

            schedule = ScheduledTime.ToString("dd/MM/yyyy hh:mm:ss tt")

            'Update Schedular
            ChangeSchedular()

            'Execute task
            'If executeTask Then PRINCIPAL()

            'Save Log
            If GuardaLog Then GuardaLogBitacora(IIf(executeTask, "Proceso ejecutado! ", String.Empty) + "Próxima ejecución aproximada: " + schedule)

        Catch ex As Exception
            ChangeSchedular(Date.Now.AddHours(1))
            GuardaLogBitacora(ex.Message)
        End Try
    End Sub

    Private Sub SchedularCallback(e As Object)
        ScheduleService()
    End Sub

    Private Sub ChangeSchedular(Optional ByVal scheduledTimeTemp As Date = Nothing)
        scheduledTimeTemp = IIf(scheduledTimeTemp = Date.MinValue, ScheduledTime, scheduledTimeTemp)

        'Get the difference in Minutes between the Scheduled and Current Time.
        Dim timeSpan As TimeSpan = scheduledTimeTemp.Subtract(Date.Now)
        Dim dueTime As Integer = Convert.ToInt32(timeSpan.TotalMilliseconds)

        'Change the Timer's Due Time.
        Schedular.Change(dueTime, Timeout.Infinite)
    End Sub
    'FIN DEL TIMER

    Private Sub GuardaLogBitacora(ByVal text As String)
        Dim path As String = ConfigurationManager.AppSettings("PathLog")

        If Not Directory.Exists(path) Then path = "C:\"

        Dim rutalog As String = IO.Path.Combine(path, "Log_BitacoraProcesarCorreosG.txt")

        File.AppendAllText(rutalog, Date.Now().ToString("dd/MM/yyyy hh:mm:ss tt") + " : " + text + vbCrLf + vbCrLf)
    End Sub

End Class