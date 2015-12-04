using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using EntityCreator.Properties;

namespace EntityCreator
{
    public partial class MainForm : Form
    {
        private readonly List<string> messagesOfTheDay = new List<string>
        {
            "What a lovely day",
            "You look very nice today",
            "You deserve everything",
            "Everything is for you",
            "You can do it",
            "Just do it",
            "Make your wishes come true"
        };

        public MainForm()
        {
            InitializeComponent();
            InitializeConfiguration();
        }

        private void InitializeConfiguration()
        {
            var random = new Random();
            var index = random.Next(messagesOfTheDay.Count);
            statusLabel.Text = messagesOfTheDay[index];
            DefaultConfiguration.IsMultiThreadSupport = false;
            DefaultConfiguration.ThrowExceptionOnNegligibleErrors = false;
        }

        private void OpenFileButton_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (!string.IsNullOrWhiteSpace(openFileDialog.FileName))
                {
                    excelFileText.Text = openFileDialog.FileName + openFileDialog.SafeFileName;
                }
            }
        }

        private void createEntitiesButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(crmServerText.Text) ||
                string.IsNullOrWhiteSpace(usernameText.Text) ||
                string.IsNullOrWhiteSpace(passwordText.Text) ||
                string.IsNullOrWhiteSpace(domainText.Text))
            {
                MessageBox.Show(Resources.MainForm_createEntitiesButton_Click_Please_fill_all_fields_);
            }
            else if (string.IsNullOrWhiteSpace(openFileDialog.FileName))
            {
                MessageBox.Show(Resources.MainForm_createEntitiesButton_Click_No_file_found_in_specified_path_);
            }
            else
            {
                var stopwatch = new Stopwatch();
                var program = new CrmHandler(crmServerText.Text, domainText.Text, usernameText.Text, passwordText.Text);
                List<Exception> warnings;

                stopwatch.Start();
                var errors = program.CreateEntity(GetEntityTemplateFromFile(), out warnings);
                stopwatch.Stop();

                var outputMessage = string.Format("Completed in {0} ", stopwatch.Elapsed);
                if (warnings.Any())
                {
                    WriteWarningLog(warnings);
                    outputMessage += "\n" + string.Format("There are {0} warnings. Please see warningsLog", warnings.Count);
                }

                if (errors.Any())
                {
                    WriteErrorLog(errors);
                    outputMessage += "\n" + string.Format("There are {0} errors. Please see errorLog", errors.Count);
                }
                MessageBox.Show(outputMessage);
            }
        }

        private void WriteErrorLog(List<Exception> errors)
        {
            WriteLog(errors, "errorlog");
        }

        private void WriteWarningLog(List<Exception> warnings)
        {
            WriteLog(warnings, "warninglog");
        }

        private void WriteLog(List<Exception> issueList, string fileName)
        {
            var errorMessageLong = string.Empty;
            foreach (var exception in issueList)
            {
                errorMessageLong += exception + "\n";
            }

            var path = fileName + ".txt";
            using (var sw = new StreamWriter(path, true))
            {
                sw.WriteLine(errorMessageLong);
                sw.Close();
            }
        }

        private EntityTemplate GetEntityTemplateFromFile()
        {
            var excelHelper = new ExcedHelper(openFileDialog.FileName);
            return excelHelper.GetEntityTemplateFromFile();
        }

        private void exportSampleButton_Click(object sender, EventArgs e)
        {
            var local = Environment.CurrentDirectory;
            ExtractEmbeddedResource(local, "EntityCreator", "Template.xlsx");
        }

        private static void ExtractEmbeddedResource(string outputDir, string resourceLocation, string file)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceLocation + @"." + file))
            {
                if (stream != null)
                {
                    using (var fileStream = new FileStream(Path.Combine(outputDir, file), FileMode.Create))
                    {
                        for (var i = 0; i < stream.Length; i++)
                        {
                            fileStream.WriteByte((byte) stream.ReadByte());
                        }
                        fileStream.Close();
                    }
                }
            }
        }
    }
}