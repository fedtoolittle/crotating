using Crotating.Services;
using Crotating.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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

            btnRun.Enabled = hasSourceFile;

            lblStatus.Text = hasSourceFile
                ? "Ready to process."
                : "Please select a source Excel file.";
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
                var reader = new ExcelReader();
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
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Processing Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                lblStatus.Text = "Processing failed.";
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            UpdateUiState();
        }
    }
}
