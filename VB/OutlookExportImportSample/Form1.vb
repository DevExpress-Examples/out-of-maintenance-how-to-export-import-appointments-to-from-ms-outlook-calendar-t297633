Imports DevExpress.XtraScheduler
Imports DevExpress.XtraScheduler.Outlook
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace OutlookExportImportSample
    Partial Public Class Form1
        Inherits DevExpress.XtraBars.Ribbon.RibbonForm

        Public Sub New()
            InitializeComponent()
            Me.ribbonControl1.SelectedPage = Me.ribbonPage1
            Me.schedulerStorage1.Appointments.ResourceSharing = True
            Me.schedulerControl1.GroupType = SchedulerGroupType.Resource
        End Sub

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
            ' TODO: This line of code loads data into the 'scheduleTestDataSet.Resources' table. You can move, or remove it, as needed.
            Me.resourcesTableAdapter.Fill(Me.scheduleTestDataSet.Resources)
            ' TODO: This line of code loads data into the 'scheduleTestDataSet.Appointments' table. You can move, or remove it, as needed.
            Me.appointmentsTableAdapter.Fill(Me.scheduleTestDataSet.Appointments)
        End Sub
        Private Sub OnAppointmentChangedInsertedDeleted(ByVal sender As Object, ByVal e As PersistentObjectsEventArgs) Handles schedulerStorage1.AppointmentsInserted, schedulerStorage1.AppointmentsChanged, schedulerStorage1.AppointmentsDeleted
            appointmentsTableAdapter.Update(scheduleTestDataSet)
            scheduleTestDataSet.AcceptChanges()
        End Sub

        Private Sub barbtnOutlookExportSingle_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles barbtnOutlookExportSelected.ItemClick
'            #Region "#ExportSelectedAppointments"
            Dim apts As AppointmentBaseCollection = schedulerControl1.SelectedAppointments
            If apts IsNot Nothing Then
                Dim storage As New SchedulerStorage()
                For Each apt As Appointment In apts
                    Dim aptCopy As Appointment = apt.Copy()
                    storage.Appointments.Add(aptCopy)
                Next apt
                storage.ExportToOutlook()
            End If
'            #End Region ' #ExportSelectedAppointments
        End Sub

        Private Sub barbtnExportUsingCriteria_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles barbtnExportUsingCriteria.ItemClick
            lbcLog.Items.Clear()
            ExportUsingCriteria()
        End Sub
        #Region "#ExportUsingCriteria"
        Private Sub ExportUsingCriteria()
            Dim exporter As OutlookExport = TryCast(schedulerControl1.Storage.CreateOutlookExporter(), OutlookExport)
            If exporter IsNot Nothing Then
                AddHandler exporter.AppointmentExporting, AddressOf exporter_AppointmentExporting
                AddHandler exporter.AppointmentExported, AddressOf exporter_AppointmentExported
                AddHandler exporter.OnException, AddressOf exporter_OnException
                exporter.CalendarFolderName = OutlookExchangeHelper.GetOutlookCalendarFolders().FirstOrDefault().FullPath
                Using stream As New MemoryStream()
                    exporter.Export(stream)
                End Using
            End If
        End Sub
        Private Sub exporter_AppointmentExporting(ByVal sender As Object, ByVal e As AppointmentExportingEventArgs)
            AddToLog(String.Format("Exporting Subj:{0}, started at {1:F} ...", e.Appointment.Subject, e.Appointment.Start))
            If e.Appointment.IsRecurring Then
                e.Cancel = True
                AddToLog("Cancelled because of its type (recurring).")
            End If
        End Sub
        Private Sub exporter_AppointmentExported(ByVal sender As Object, ByVal e As AppointmentExportedEventArgs)
            AddToLog(String.Format("Successfully exported Subj:{0}, started at {1:F}!", e.Appointment.Subject, e.Appointment.Start))
        End Sub
        Private Sub exporter_OnException(ByVal sender As Object, ByVal e As ExchangeExceptionEventArgs)
            Dim errText As String = e.OriginalException.Message
            AddToLog(errText)
            Dim exporter As OutlookExport = DirectCast(sender, OutlookExport)
            exporter.Terminate()
            e.Handled = True
            'throw e.OriginalException;
        End Sub
        #End Region ' #ExportUsingCriteria

        Private Sub AddToLog(ByVal logText As String)
            lbcLog.Items.Add(logText)
        End Sub

        Private Sub barbtnOutlookImport_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles barbtnOutlookImport.ItemClick
            lbcLog.Items.Clear()
            ImportUsingCriteria()
        End Sub

        #Region "#ImportUsingCriteria"
        Private Sub ImportUsingCriteria()
            Dim importer As OutlookImport = TryCast(schedulerControl1.Storage.CreateOutlookImporter(), OutlookImport)
            If importer IsNot Nothing Then
                AddHandler importer.AppointmentImporting, AddressOf importer_AppointmentImporting
                AddHandler importer.AppointmentImported, AddressOf importer_AppointmentImported
                AddHandler importer.OnException, AddressOf importer_OnException
                importer.CalendarFolderName = OutlookExchangeHelper.GetOutlookCalendarFolders().FirstOrDefault().FullPath
                Using stream As New MemoryStream()
                    importer.Import(stream)
                End Using
            End If
        End Sub
        Private Sub importer_AppointmentImporting(ByVal sender As Object, ByVal e As AppointmentImportingEventArgs)
            Dim args As OutlookAppointmentImportingEventArgs = TryCast(e, OutlookAppointmentImportingEventArgs)
            AddToLog(String.Format("Importing Subj:{0}, started at {1:F} ...", args.OutlookAppointment.Subject, args.OutlookAppointment.Start))
            If args.OutlookAppointment.BusyStatus = DevExpress.XtraScheduler.Outlook.Interop.OlBusyStatus.olWorkingElsewhere Then
                e.Cancel = True
                AddToLog("Cancelled because of its busy type (working elsewhere).")
            End If
        End Sub
        Private Sub importer_AppointmentImported(ByVal sender As Object, ByVal e As AppointmentImportedEventArgs)
            Dim args As OutlookAppointmentImportedEventArgs = TryCast(e, OutlookAppointmentImportedEventArgs)
            AddToLog(String.Format("Successfully imported Subj:{0}, started at {1:F}!", args.OutlookAppointment.Subject, args.OutlookAppointment.Start))
        End Sub

        Private Sub importer_OnException(ByVal sender As Object, ByVal e As ExchangeExceptionEventArgs)
            Dim errText As String = e.OriginalException.Message
            AddToLog(errText)
            Dim importer As OutlookImport = DirectCast(sender, OutlookImport)
            importer.Terminate()
            e.Handled = True
            'throw e.OriginalException;
        End Sub
        #End Region ' #ImportUsingCriteria
    End Class
End Namespace
