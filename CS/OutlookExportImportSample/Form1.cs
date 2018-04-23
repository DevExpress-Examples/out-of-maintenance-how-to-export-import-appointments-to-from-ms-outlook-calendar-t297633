using DevExpress.XtraScheduler;
using DevExpress.XtraScheduler.Outlook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookExportImportSample {
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm {
        public Form1() {
            InitializeComponent();
            this.ribbonControl1.SelectedPage = this.ribbonPage1;
            this.schedulerStorage1.Appointments.ResourceSharing = true;
            this.schedulerControl1.GroupType = SchedulerGroupType.Resource;
        }

        private void Form1_Load(object sender, EventArgs e) {
            // TODO: This line of code loads data into the 'scheduleTestDataSet.Resources' table. You can move, or remove it, as needed.
            this.resourcesTableAdapter.Fill(this.scheduleTestDataSet.Resources);
            // TODO: This line of code loads data into the 'scheduleTestDataSet.Appointments' table. You can move, or remove it, as needed.
            this.appointmentsTableAdapter.Fill(this.scheduleTestDataSet.Appointments);
        }
        private void OnAppointmentChangedInsertedDeleted(object sender, PersistentObjectsEventArgs e) {
            appointmentsTableAdapter.Update(scheduleTestDataSet);
            scheduleTestDataSet.AcceptChanges();
        }

        private void barbtnOutlookExportSingle_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            #region #ExportSelectedAppointments
            AppointmentBaseCollection apts = schedulerControl1.SelectedAppointments;
            if (apts != null) {
                SchedulerStorage storage = new SchedulerStorage();
                foreach (Appointment apt in apts) {
                    Appointment aptCopy = apt.Copy();
                    storage.Appointments.Add(aptCopy);
                }
                storage.ExportToOutlook();
            }
            #endregion #ExportSelectedAppointments
        }

        private void barbtnExportUsingCriteria_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            lbcLog.Items.Clear();
            ExportUsingCriteria();
        }
        #region #ExportUsingCriteria
        private void ExportUsingCriteria() {
            OutlookExport exporter = schedulerControl1.Storage.CreateOutlookExporter() as OutlookExport;
            if (exporter != null) {
                exporter.AppointmentExporting += exporter_AppointmentExporting;
                exporter.AppointmentExported += exporter_AppointmentExported;
                exporter.OnException += exporter_OnException;
                exporter.CalendarFolderName = OutlookExchangeHelper.GetOutlookCalendarFolders().FirstOrDefault().FullPath;
                using (MemoryStream stream = new MemoryStream()) {
                    exporter.Export(stream);
                }
            }
        }
        void exporter_AppointmentExporting(object sender, AppointmentExportingEventArgs e) {
            AddToLog(String.Format("Exporting Subj:{0}, started at {1:F} ...", e.Appointment.Subject, e.Appointment.Start));
            if (e.Appointment.IsRecurring) {
                e.Cancel = true;
                AddToLog("Cancelled because of its type (recurring).");
            }
        }
        void exporter_AppointmentExported(object sender, AppointmentExportedEventArgs e) {
            AddToLog(String.Format("Successfully exported Subj:{0}, started at {1:F}!", e.Appointment.Subject, e.Appointment.Start)); 
        }
        void exporter_OnException(object sender, ExchangeExceptionEventArgs e) {
            string errText = e.OriginalException.Message;
            AddToLog(errText);
            OutlookExport exporter = (OutlookExport)sender;
            exporter.Terminate();
            e.Handled = true;
            //throw e.OriginalException;
        }
        #endregion #ExportUsingCriteria

        private void AddToLog(string logText) {
            lbcLog.Items.Add(logText);
        }

        private void barbtnOutlookImport_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            lbcLog.Items.Clear();
            ImportUsingCriteria();
        }

        #region #ImportUsingCriteria
        private void ImportUsingCriteria() {
            OutlookImport importer = schedulerControl1.Storage.CreateOutlookImporter() as OutlookImport;
            if (importer != null) {
                importer.AppointmentImporting += importer_AppointmentImporting;
                importer.AppointmentImported += importer_AppointmentImported;
                importer.OnException += importer_OnException;
                importer.CalendarFolderName = OutlookExchangeHelper.GetOutlookCalendarFolders().FirstOrDefault().FullPath;
                using (MemoryStream stream = new MemoryStream()) {
                    importer.Import(stream);
                }
            }
        }
        void importer_AppointmentImporting(object sender, AppointmentImportingEventArgs e) {
            OutlookAppointmentImportingEventArgs args = e as OutlookAppointmentImportingEventArgs;
            AddToLog(String.Format("Importing Subj:{0}, started at {1:F} ...", args.OutlookAppointment.Subject, args.OutlookAppointment.Start));
            if (args.OutlookAppointment.BusyStatus == DevExpress.XtraScheduler.Outlook.Interop.OlBusyStatus.olWorkingElsewhere) {
                e.Cancel = true;
                AddToLog("Cancelled because of its busy type (working elsewhere).");
            }
        }
        void importer_AppointmentImported(object sender, AppointmentImportedEventArgs e) {
            OutlookAppointmentImportedEventArgs args = e as OutlookAppointmentImportedEventArgs;
            AddToLog(String.Format("Successfully imported Subj:{0}, started at {1:F}!", args.OutlookAppointment.Subject, args.OutlookAppointment.Start)); 
        }

        void importer_OnException(object sender, ExchangeExceptionEventArgs e) {
            string errText = e.OriginalException.Message;
            AddToLog(errText);
            OutlookImport importer = (OutlookImport)sender;
            importer.Terminate();
            e.Handled = true;
            //throw e.OriginalException;
        }
        #endregion #ImportUsingCriteria
    }
}
