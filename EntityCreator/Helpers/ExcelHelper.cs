using System;
using System.Collections.Generic;
using System.Linq;
using EntityCreator.Models;
using Microsoft.Office.Interop.Excel;
using Microsoft.Xrm.Sdk;

namespace EntityCreator.Helpers
{
    public sealed class ExcelHelper : IDisposable
    {
        #region DefaultExcelVariableDefinations

        private const int defaultOptionSetStartIndex = 0;

        private const int defaultEntitySheetIndex = 1;
        private const int defaultWebresourceSheetIndex = 2;

        private const int entityTemplateDefinationRow = 2;
        private const int attributeTemplateFirstRow = 4;
        private const int webresourceTemplateFirstRow = 2;

        private const int entityLogicalNameColumn = 1;
        private const int entityDisplayNameColumn = 2;
        private const int entityDescriptionColumn = 3;
        private const int entityDisplayNamePluralColumn = 4;

        private const int attributeLogicalNameColumn = 1;
        private const int attributeDisplayNameColumn = 2;
        private const int attributeDescriptionColumn = 3;
        private const int attributeAttributeTypeColumn = 4;
        private const int attributeLookupEntityLogicalNameColumn = 5;
        private const int attributeOptionSetListColumn = 6;
        private const int attributeGlobalOptionSetListLogicalNameColumn = 7;
        private const int attributeMinValueColumn = 8;
        private const int attributeMaxValueColumn = 9;
        private const int attributeRequiredColumn = 10;
        private const int attributeOtherDisplayNameColumn = 11;
        private const int attributeOtherDescriptionColumn = 12;

        private const int webresourceLogicalNameColumn = 1;
        private const int webresourceDisplayNameColumn = 2;
        private const int webresourceDescriptionColumn = 3;
        private const int webresourceTypeColumn = 4;
        private const int webresourceContentColumn = 5;

        #endregion

        private readonly Application xlsApplication;
        private readonly Workbook xlsWorkbook;

        public ExcelHelper(string filePath)
        {
            xlsApplication = new Application();
            xlsWorkbook = xlsApplication.Workbooks.Open(filePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t",
                false, false, 0, true, 1, 0);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public EntityTemplate GetEntityTemplateFromFile(string excelFile)
        {
            var entityTemplate = new EntityTemplate {AttributeList = new List<AttributeTemplate>(),Warnings = new List<Exception>(),Errors = new List<Exception>()};
            var xlsWorksheet = (Worksheet) xlsWorkbook.Worksheets.Item[defaultEntitySheetIndex];
            
            var xlsRange = xlsWorksheet.UsedRange;

            if (GetEntityTemplate(excelFile, entityTemplate.Errors, entityTemplate, xlsRange))
            {
                entityTemplate.WillCreateEntity = false;
            }
            else
            {
                entityTemplate.WillCreateEntity = true;
            }
            for (var currentRow = attributeTemplateFirstRow; currentRow <= xlsRange.Rows.Count; currentRow++)
            {
                var logicalName = GetCellValueAsString(xlsRange, currentRow, attributeLogicalNameColumn);
                var displayName = GetCellValueAsString(xlsRange, currentRow, attributeDisplayNameColumn);
                if (string.IsNullOrWhiteSpace(logicalName) || string.IsNullOrWhiteSpace(displayName))
                {
                    //entityTemplate.Errors.Add(
                    //    new Exception(string.Format("Attribute LogicalName or DisplayName can not be empty. File: {0}",
                    //        excelFile)));
                    continue;
                }

                var attributeTemplate = GetAttributeTemplate(xlsRange, currentRow, logicalName, excelFile, entityTemplate.Errors,
                    entityTemplate.Warnings);
                if (attributeTemplate == null)
                {
                    continue;
                }

                if (ApplyFileSpecificOperations(excelFile, entityTemplate.Errors, attributeTemplate, xlsRange, currentRow))
                {
                    continue;
                }

                entityTemplate.AttributeList.Add(attributeTemplate);
            }
            if (xlsWorkbook.Worksheets.Count > 1)
            {
                var webresourceSheet = (Worksheet)xlsWorkbook.Worksheets.Item[defaultWebresourceSheetIndex];
                CreateWebResources(webresourceSheet, entityTemplate, excelFile, entityTemplate.Errors, entityTemplate.Warnings);
            }
            
            return entityTemplate;
        }

        private bool ApplyFileSpecificOperations(string excelFile, List<Exception> errorList,
            AttributeTemplate attributeTemplate,
            Range xlsRange, int currentRow)
        {
            if (attributeTemplate.AttributeType == typeof (Lookup))
            {
                attributeTemplate.LookupEntityLogicalName = GetCellValueAsString(xlsRange, currentRow,
                    attributeLookupEntityLogicalNameColumn);
                if (string.IsNullOrWhiteSpace(attributeTemplate.LookupEntityLogicalName))
                {
                    errorList.Add(
                        new Exception(
                            string.Format(
                                "Attribute LookupEntityLogicalName can not be empty for LookUp fields. File: {0}",
                                excelFile)));
                    return true;
                }

                attributeTemplate.LookupEntityLogicalName = attributeTemplate.LookupEntityLogicalName.Trim();
            }
            else if (attributeTemplate.AttributeType == typeof (OptionSet))
            {
                attributeTemplate.OptionSetList =
                    GetOptionSetList(GetCellValueAsString(xlsRange, currentRow, attributeOptionSetListColumn));
            }
            else if (attributeTemplate.AttributeType == typeof (bool))
            {
                attributeTemplate.OptionSetList =
                    GetOptionSetList(GetCellValueAsString(xlsRange, currentRow, attributeOptionSetListColumn));
            }
            else if (attributeTemplate.AttributeType == typeof (GlobalOptionSet))
            {
                attributeTemplate.OptionSetList =
                    GetOptionSetList(GetCellValueAsString(xlsRange, currentRow, attributeOptionSetListColumn));
                attributeTemplate.GlobalOptionSetListLogicalName = GetCellValueAsString(xlsRange, currentRow,
                    attributeGlobalOptionSetListLogicalNameColumn);
                if (string.IsNullOrWhiteSpace(attributeTemplate.GlobalOptionSetListLogicalName))
                {
                    errorList.Add(
                        new Exception(
                            string.Format(
                                "Attribute GlobalOptionSetListLogicalName can not be empty for GlobalOptionSet fields. File: {0}",
                                excelFile)));
                    return true;
                }
                attributeTemplate.GlobalOptionSetListLogicalName =
                    attributeTemplate.GlobalOptionSetListLogicalName.Trim();
            }
            else if (attributeTemplate.AttributeType == typeof(NNRelation))
            {
                attributeTemplate.LookupEntityLogicalName = GetCellValueAsString(xlsRange, currentRow,
                    attributeLookupEntityLogicalNameColumn);
                if (string.IsNullOrWhiteSpace(attributeTemplate.LookupEntityLogicalName))
                {
                    errorList.Add(
                        new Exception(
                            string.Format(
                                "Attribute LookupEntityLogicalName can not be empty for NN fields. File: {0}",
                                excelFile)));
                    return true;
                }
            }
            return false;
        }

        private AttributeTemplate GetAttributeTemplate(Range xlsRange, int currentRow, string logicalName,
            string excelFile, List<Exception> errorList, List<Exception> warningList)
        {
            var displayName = GetCellValueAsString(xlsRange, currentRow, attributeDisplayNameColumn);
            var displayNameShort = displayName.Length > DefaultConfiguration.AttributeDisplayNameMaxLength
                ? displayName.Substring(0, DefaultConfiguration.AttributeDisplayNameMaxLength)
                : displayName;
            var minLength = default(int);
            var maxLength = default(int);
            var minLengthStr = GetCellValueAsString(xlsRange, currentRow, attributeMinValueColumn);
            var maxLengthStr = GetCellValueAsString(xlsRange, currentRow, attributeMaxValueColumn);
            maxLengthStr = maxLengthStr.Replace(".", "").Replace(",",".");
            var isRequiredStr = GetCellValueAsString(xlsRange, currentRow, attributeRequiredColumn).ToLower();
            var isRequired = false;// DefaultConfiguration.YesKeywordList.Contains(isRequiredStr);

            if (!string.IsNullOrWhiteSpace(maxLengthStr) && !int.TryParse(maxLengthStr, out maxLength))
            {
                maxLength = default(int);
            }
            minLengthStr = minLengthStr.Replace(".", "").Replace(",",".");
            

            if (!string.IsNullOrWhiteSpace(minLengthStr) && !int.TryParse(minLengthStr, out minLength))
            {
                minLength = default(int);
            }
            var attributeType =
                CommonHelper.GetAttributeType(GetCellValueAsString(xlsRange, currentRow, attributeAttributeTypeColumn),
                    excelFile, errorList, warningList);
            if (attributeType == typeof (Exception))
            {
                return null;
            }

            var attributeTemplate = new AttributeTemplate
            {
                LogicalName = logicalName,
                DisplayName = displayName,
                DisplayNameShort = displayNameShort,
                Description = GetCellValueAsString(xlsRange, currentRow, attributeDescriptionColumn),
                MinLength = minLength,
                MaxLength = maxLength,
                AttributeType = attributeType,
                IsRequired = isRequired,
                OtherDisplayName = GetCellValueAsString(xlsRange, currentRow, attributeOtherDisplayNameColumn),
                OtherDescription = GetCellValueAsString(xlsRange, currentRow, attributeOtherDescriptionColumn)
            };

            attributeTemplate.LogicalName = attributeTemplate.LogicalName.Trim();
            attributeTemplate.DisplayName = attributeTemplate.DisplayName.Trim();

            return attributeTemplate;
        }

        private static bool GetEntityTemplate(string excelFile, List<Exception> error, EntityTemplate entityTemplate,
            Range xlsRange)
        {
            entityTemplate.LogicalName = GetCellValueAsString(xlsRange, entityTemplateDefinationRow,
                entityLogicalNameColumn);
            entityTemplate.DisplayName = GetCellValueAsString(xlsRange, entityTemplateDefinationRow,
                entityDisplayNameColumn);
            entityTemplate.Description = GetCellValueAsString(xlsRange, entityTemplateDefinationRow,
                entityDescriptionColumn);
            entityTemplate.DisplayNamePlural = GetCellValueAsString(xlsRange, entityTemplateDefinationRow,
                entityDisplayNamePluralColumn);

            if (string.IsNullOrWhiteSpace(entityTemplate.LogicalName) ||
                string.IsNullOrWhiteSpace(entityTemplate.DisplayName) ||
                string.IsNullOrWhiteSpace(entityTemplate.DisplayNamePlural))
            {
                error.Add(
                    new Exception(
                        string.Format(
                            "Entity LogicalName, DisplayName or DisplayNamePlural can not be empty. File: {0}",
                            excelFile)));
                return true;
            }

            entityTemplate.LogicalName = entityTemplate.LogicalName.ToLower().Trim();
            entityTemplate.DisplayName = entityTemplate.DisplayName.Trim();
            entityTemplate.DisplayNamePlural = entityTemplate.DisplayName.Trim();
            return false;
        }

        private void CreateWebResources(Worksheet webresourceSheet, EntityTemplate entityTemplate, string excelFile,
            List<Exception> error, List<Exception> warning)
        {
            var xlsRange = webresourceSheet.UsedRange;
            for (var currentRow = webresourceTemplateFirstRow; currentRow < xlsRange.Rows.Count; currentRow++)
            {
                var logicalName = GetCellValueAsString(xlsRange, currentRow, webresourceLogicalNameColumn).Replace("\n", "");
                var displayName = GetCellValueAsString(xlsRange, currentRow, webresourceDisplayNameColumn).Replace("\n", "");
                var description = GetCellValueAsString(xlsRange, currentRow, webresourceDescriptionColumn).Replace("\n", "");
                var type = GetCellValueAsString(xlsRange, currentRow, webresourceTypeColumn).Replace("\n", "");
                var content =
                    CommonHelper.EncodeTo64(GetCellValueAsString(xlsRange, currentRow, webresourceContentColumn));

                var resourceType = CommonHelper.GetWebResourceType(type, excelFile, error, warning);
                if (resourceType < 0)
                {
                    continue;
                }

                var webResource = new WebResource
                {
                    Content = content,
                    DisplayName = displayName,
                    Description = description,
                    Name = logicalName,
                    LogicalName = WebResource.EntityLogicalName,
                    WebResourceType = new OptionSetValue(resourceType)
                };
                entityTemplate.WebResource = entityTemplate.WebResource ?? new List<WebResource>();
                entityTemplate.WebResource.Add(webResource);
            }
        }

        private static string GetCellValueAsString(Range xlsRange, int row, int column)
        {
            var cell = (xlsRange.Cells[row, column] as Range);
            var cellValue = cell == null ? null : cell.Value2;
            return cellValue == null ? string.Empty : cellValue.ToString();
        }

        private List<OptionSetTemplate> GetOptionSetList(string optionSetListString)
        {
            var optionSetTemplateList = new List<OptionSetTemplate>();
            if (!string.IsNullOrWhiteSpace(optionSetListString))
            {
                var optionSetArray = optionSetListString.Split(DefaultConfiguration.OptionSetSplicChar).ToList();
                var startIndex = defaultOptionSetStartIndex;
                foreach (var optionSetValueStrLoop in optionSetArray)
                {
                    var optionSetValueStr = optionSetValueStrLoop.Trim();
                    var labelValueList = optionSetValueStr.Split('=').ToList();
                    var optionSetTemplate = new OptionSetTemplate
                    {
                        Label = labelValueList.First().Trim(),
                        Value = Convert.ToInt32(labelValueList.Last().Trim())
                    };
                    optionSetTemplateList.Add(optionSetTemplate);
                    startIndex++;
                }
            }

            return optionSetTemplateList;
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                xlsWorkbook.Close(true, null, null);
                xlsApplication.Quit();
            }
        }
    }
}