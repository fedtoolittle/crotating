using Crotating.Models;
using Crotating.Services;
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

namespace Crotating
{
    public partial class Form1 : Form
    {
        private List<WorkEntry> _loadedEntries;

        public Form1()
        {

            InitializeComponent();
            cmbInputFormat.Items.Add("Crabal Time SHEET");
            cmbInputFormat.Items.Add("Crabal Time CARD");
            cmbInputFormat.SelectedIndex = 0;

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (_loadedEntries == null || _loadedEntries.Count == 0)
            {
                MessageBox.Show(
                    "No data available to export.",
                    "Export Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                dialog.Title = "Save exported summary";
                dialog.FileName = "WorkSummary.xlsx";

                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                var exporter = new ExcelExporter();
                exporter.ExportSummary(_loadedEntries, dialog.FileName);

                MessageBox.Show(
                    "Export completed successfully.",
                    "Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private void btnBrowseSource_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                dialog.Title = "Select source Excel file";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtSourceFile.Text = dialog.FileName;
                    UpdateUiState();
                }
            }
        }


        private void UpdateUiState()
        {
            bool hasSourceFile =
        !string.IsNullOrWhiteSpace(txtSourceFile.Text) &&
        System.IO.File.Exists(txtSourceFile.Text);

            bool hasFormat =
                cmbInputFormat.SelectedIndex >= 0;

            btnRun.Enabled = hasSourceFile && hasFormat;

            lblStatus.Text = btnRun.Enabled
                ? "Ready to process."
                : "Please select a source file and format.";
        }

        //private void btnRun_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (!System.IO.File.Exists(txtSourceFile.Text))
        //        {
        //            MessageBox.Show(
        //                "The selected source file does not exist.",
        //                "Invalid File",
        //                MessageBoxButtons.OK,
        //                MessageBoxIcon.Error);
        //
        //            UpdateUiState();
        //            return;
        //        }
        //
        //        lblStatus.Text = "Validation successful. Processing will be added next.";
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(
        //            ex.Message,
        //            "Unexpected Error",
        //            MessageBoxButtons.OK,
        //            MessageBoxIcon.Error);
        //
        //        lblStatus.Text = "Error occurred.";
        //    }
        //}

        private void btnRun_Click(object sender, EventArgs e)
        {
            try
            {
                IWorkEntryReader reader;

                switch (cmbInputFormat.SelectedIndex)
                {
                    case 0:
                        reader = new CrabalTimesheetReader();
                        break;

                    case 1:
                        reader = new CrabalTimecardReader();
                        break;

                    default:
                        throw new InvalidOperationException("Invalid input format selected.");
                }

                _loadedEntries = reader.ReadEntries(txtSourceFile.Text);

                var aggregator = new WorkAggregator();
                var summaries = aggregator.AggregateByPersonAndDay(_loadedEntries);

                var service = new WorkSummaryService();
                var table = service.BuildExportTable(summaries);

                //foreach (var s in summaries)
                //{
                //    System.Diagnostics.Debug.WriteLine(
                //        s.Name +   " | " + s.Date.ToShortDateString() + " | " + s.TotalHours + " hours");
                //}



                lblStatus.Text = "Loaded " + _loadedEntries.Count + " rows successfully.";
                btnExport.Enabled = true;

            }
            catch (InvalidDataException ex)
            {
                MessageBox.Show(
                    "The selected file does not match the chosen format.\n\n" +
                    ex.Message,
                    "Invalid File Format",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "An unexpected error occurred.\n\n" + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            UpdateUiState();
        }

        private void cmbInputFormat_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateUiState();
        }
    }
}
