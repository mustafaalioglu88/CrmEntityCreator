﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using EntityCreator.Helpers;
using EntityCreator.Models;
using EntityCreator.Properties;

namespace EntityCreator
{
    public partial class MainForm : Form
    {
        private readonly Stopwatch stopwatch = new Stopwatch();
        //private Action<int> calculateCreatedItems;
        //private Action<string> setMessageOfTheDayAction;
        private Dictionary<string, EntityTemplate> entityTemplates;
        //private int numberOfItemsCreated;
        //private int totalNumberOfItemsWillBeCreated;

        public MainForm()
        {
            InitializeComponent();
            InitializeConfiguration();
        }

        //private void CalculateNumberOfItemsCreated(int numberOfItemsCreated)
        //{
        //    this.numberOfItemsCreated += numberOfItemsCreated;
        //    var percentage = (100*this.numberOfItemsCreated)/totalNumberOfItemsWillBeCreated;
        //    Invoke((MethodInvoker) delegate
        //    {
        //        var value = percentage - progressBar.Value;
        //        progressBar.Increment(value); // runs on UI thread
        //    });
        //}

        private void SetMessageOfTheDayUi(string messageOfTheDay)
        {
            Invoke((MethodInvoker)delegate
            {
                statusLabel.Text = messageOfTheDay;
            });
        }

        private void InitializeConfiguration()
        {
            this.Text += " v." + System.Reflection.Assembly.GetExecutingAssembly()
                                           .GetName()
                                           .Version; 
            SetMessageOfTheDay();
            //calculateCreatedItems += CalculateNumberOfItemsCreated;
            //setMessageOfTheDayAction += SetMessageOfTheDayUi;
        }

        private void SetMessageOfTheDay()
        {
            var random = new Random();
            var index = random.Next(DefaultConfiguration.MessagesOfTheDay.Count);
            statusLabel.Text = DefaultConfiguration.MessagesOfTheDay[index];
        }

        private void OpenFileButton_Click(object sender, EventArgs e)
        {
            if (openFolderBrowserDialog.ShowDialog() == DialogResult.OK &&
                !string.IsNullOrWhiteSpace(openFolderBrowserDialog.SelectedPath))
            {
                excelFileText.Text = openFolderBrowserDialog.SelectedPath;
            }
        }

        private void createEntitiesButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(crmServerText.Text) || string.IsNullOrWhiteSpace(usernameText.Text) ||
                string.IsNullOrWhiteSpace(passwordText.Text) || string.IsNullOrWhiteSpace(domainText.Text))
            {
                MessageBox.Show(Resources.MainForm_createEntitiesButton_Click_Please_fill_all_fields_);
            }
            else if (string.IsNullOrWhiteSpace(openFolderBrowserDialog.SelectedPath))
            {
                MessageBox.Show(Resources.MainForm_createEntitiesButton_Click_No_file_found_in_specified_path_);
            }
            else
            {
                createEntitiesButton.Enabled = false;
                //progressBar.Maximum = 100;
                //progressBar.Step = 1;
                //progressBar.Value = 0;
                var thread = new Thread(StartExectionThread);
                thread.Start();
            }
        }

        private string StartExecution(Action<string> setMessageOfTheDayLabel)
        {
            setMessageOfTheDayLabel("Execution started");
            var program = new CrmHelper(crmServerText.Text, domainText.Text, usernameText.Text, passwordText.Text);
            setMessageOfTheDayLabel("CRM connection created");
            stopwatch.Start();

            setMessageOfTheDayLabel("Excel files started to read");
            var excelFiles = FileHelpers.GetExcelFiles(openFolderBrowserDialog.SelectedPath);
            var genericErrorList = new Dictionary<string, List<Exception>>();
            var genericWarningList = new Dictionary<string, List<Exception>>();
            entityTemplates = new Dictionary<string, EntityTemplate>();
            int i = 1;
            foreach (var excelFile in excelFiles)
            {
                var fileName = excelFile.Split('\\').Last();
                setMessageOfTheDayLabel(i.ToString() + "/" + excelFiles.Count() + ": File reading: " + fileName);
                var entityTemplate = FileHelpers.GetEntityTemplateFromFile(excelFile);
                if (entityTemplate == null)
                    continue;
                //if (entityTemplate.WebResource != null)
                //    totalNumberOfItemsWillBeCreated += entityTemplate.AttributeList.Count +
                //                                       entityTemplate.WebResource.Count + 1;
                //else if (
                //    !entityTemplate.WillCreateEntity && entityTemplate.WebResource != null)
                //{
                //    totalNumberOfItemsWillBeCreated += entityTemplate.WebResource.Count;
                //}
                //else
                //{
                //    totalNumberOfItemsWillBeCreated += entityTemplate.AttributeList.Count + 1;
                //}
                entityTemplates.Add(excelFile, entityTemplate);
                i++;
            }

            i = 1;
            setMessageOfTheDayLabel("Excel files read has finished.");
            foreach (var template in entityTemplates)
            {
                var fileName = template.Key.Split('\\').Last();
                setMessageOfTheDayLabel(i.ToString() + "/" + entityTemplates.Count() + ": On CRM: " + fileName);
                program.CreateEntity(template.Key, template.Value, setMessageOfTheDayLabel);

                genericErrorList.Add(template.Key, template.Value.Errors);
                genericWarningList.Add(template.Key, template.Value.Warnings);

                FileHelpers.MarkFileAsProcessed(template.Key);
                i++;
            }

            stopwatch.Stop();
            var outputMessage = GenerateOutputMessage(genericWarningList, genericErrorList);
            program.Dispose();
            setMessageOfTheDayLabel("Done. Are you happy now? You can start over.");
            return outputMessage;
        }

        private string GenerateOutputMessage(Dictionary<string, List<Exception>> warnings,
            Dictionary<string, List<Exception>> errors)
        {
            var outputMessage = string.Format("Completed in {0} ", stopwatch.Elapsed);
            stopwatch.Reset();

            if (warnings.Any())
            {
                FileHelpers.WriteWarningLog(warnings);
                outputMessage += "\n" + string.Format("There are {0} warnings. Please see warningsLogs", warnings.Count);
            }

            if (errors.Any())
            {
                FileHelpers.WriteErrorLog(errors);
                outputMessage += "\n" + string.Format("There are {0} errors. Please see errorLogs", errors.Count);
            }

            return outputMessage;
        }

        private void exportSampleButton_Click(object sender, EventArgs e)
        {
            FileHelpers.ExtractResources();
        }

        private void StartExectionThread()
        {
            var outputMessage = StartExecution(SetMessageOfTheDayUi);
            MessageBox.Show(outputMessage);
        }
    }
}