using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Microsoft.Xrm.Sdk;

namespace EntityCreator
{
    public sealed class ExcedHelper : IDisposable
    {
        private const int defaultSheetIndex = 1;
        private const int defaultOptionSetStartIndex = 0;

        private const int entityTemplateRow = 2;
        private const int attributeTemplateRowStart = 4;

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

        private readonly Application xlsApplication = new Application();
        private readonly Workbook xlsWorkbook;

        public ExcedHelper(string filePath)
        {
            xlsWorkbook = xlsApplication.Workbooks.Open(filePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        }

        public EntityTemplate GetEntityTemplateFromFile()
        {
            var entityTemplate = new EntityTemplate {AttributeList = new List<AttributeTemplate>()};
            var xlsWorksheet = (Worksheet)xlsWorkbook.Worksheets.Item[defaultSheetIndex];
            var xlsRange = xlsWorksheet.UsedRange;

            entityTemplate.LogicalName = GetCellValueAsString(xlsRange, entityTemplateRow, entityLogicalNameColumn);
            entityTemplate.DisplayName = GetCellValueAsString(xlsRange, entityTemplateRow, entityDisplayNameColumn);
            entityTemplate.Description = GetCellValueAsString(xlsRange, entityTemplateRow, entityDescriptionColumn);
            entityTemplate.DisplayNamePlural = GetCellValueAsString(xlsRange, entityTemplateRow, entityDisplayNamePluralColumn);

            if (string.IsNullOrWhiteSpace(entityTemplate.LogicalName) || string.IsNullOrWhiteSpace(entityTemplate.DisplayName) || 
                string.IsNullOrWhiteSpace(entityTemplate.DisplayNamePlural))
            {
                throw new Exception("Entity LogicalName, DisplayName or DisplayNamePlural can not be empty.");
            }


            entityTemplate.LogicalName = entityTemplate.LogicalName.ToLower().Trim();
            entityTemplate.DisplayName = entityTemplate.DisplayName.Trim();
            entityTemplate.DisplayNamePlural = entityTemplate.DisplayName.Trim();
            for (var currentRow = attributeTemplateRowStart; currentRow <= xlsRange.Rows.Count; currentRow++)
            {
                var logicalName = GetCellValueAsString(xlsRange, currentRow, attributeLogicalNameColumn);
                if (string.IsNullOrWhiteSpace(logicalName))
                {
                    continue;
                }

                var minLength = default(int);
                var maxLength = default(int);
                var minLengthStr = GetCellValueAsString(xlsRange, currentRow, attributeMinValueColumn);
                var maxLengthStr = GetCellValueAsString(xlsRange, currentRow, attributeMaxValueColumn);
                if (!string.IsNullOrWhiteSpace(maxLengthStr) && int.TryParse(maxLengthStr, out maxLength))
                {
                    maxLength = default(int);
                }

                if (!string.IsNullOrWhiteSpace(minLengthStr) && int.TryParse(minLengthStr, out minLength))
                {
                    minLength = default(int);
                }

                var attributeTemplate = new AttributeTemplate
                {
                    LogicalName = logicalName,
                    DisplayName = GetCellValueAsString(xlsRange, currentRow, attributeDisplayNameColumn),
                    Description = GetCellValueAsString(xlsRange, currentRow, attributeDescriptionColumn),
                    MinLength = minLength,
                    MaxLength = maxLength,
                    AttributeType = GetAttributeType(GetCellValueAsString(xlsRange, currentRow, attributeAttributeTypeColumn))
                };
                if (string.IsNullOrWhiteSpace(attributeTemplate.LogicalName) || string.IsNullOrWhiteSpace(attributeTemplate.DisplayName))
                {
                    throw new Exception("Attribute LogicalName or DisplayName can not be empty.");
                }

                attributeTemplate.LogicalName = attributeTemplate.LogicalName.Trim();
                attributeTemplate.DisplayName = attributeTemplate.DisplayName.Trim();

                if (attributeTemplate.AttributeType == typeof (Lookup))
                {
                    attributeTemplate.LookupEntityLogicalName = GetCellValueAsString(xlsRange, currentRow, attributeLookupEntityLogicalNameColumn);
                    if (string.IsNullOrWhiteSpace(attributeTemplate.LookupEntityLogicalName))
                    {
                        throw new Exception("Attribute LookupEntityLogicalName can not be empty.");
                    }
                    attributeTemplate.LookupEntityLogicalName = attributeTemplate.LookupEntityLogicalName.Trim();
                }
                else if (attributeTemplate.AttributeType == typeof(OptionSet))
                {
                    attributeTemplate.OptionSetList = GetOptionSetList(GetCellValueAsString(xlsRange, currentRow, attributeOptionSetListColumn));
                }
                else if (attributeTemplate.AttributeType == typeof(bool))
                {
                    attributeTemplate.OptionSetList = GetOptionSetList(GetCellValueAsString(xlsRange, currentRow, attributeOptionSetListColumn));
                }
                else if (attributeTemplate.AttributeType == typeof(GlobalOptionSet))
                {
                    attributeTemplate.OptionSetList = GetOptionSetList(GetCellValueAsString(xlsRange, currentRow, attributeOptionSetListColumn));
                    attributeTemplate.GlobalOptionSetListLogicalName = GetCellValueAsString(xlsRange, currentRow, attributeGlobalOptionSetListLogicalNameColumn);
                    if (string.IsNullOrWhiteSpace(attributeTemplate.GlobalOptionSetListLogicalName))
                    {
                        throw new Exception("Attribute GlobalOptionSetListLogicalName can not be empty.");
                    }
                    attributeTemplate.GlobalOptionSetListLogicalName = attributeTemplate.GlobalOptionSetListLogicalName.Trim();
                }

                entityTemplate.AttributeList.Add(attributeTemplate);
            }

            return entityTemplate;
        }

        private static dynamic GetCellValueAsString(Range xlsRange, int row, int column)
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
                foreach (var optionSetLabel in optionSetArray)
                {
                    var optionSetTemplate = new OptionSetTemplate
                    {
                        Label = optionSetLabel.Trim(),
                        Value = startIndex
                    };
                    optionSetTemplateList.Add(optionSetTemplate);
                    startIndex++;
                }
            }

            return optionSetTemplateList;
        }

        private Type GetAttributeType(string attributeTypeString)
        {
            if (string.IsNullOrWhiteSpace(attributeTypeString))
            {
                throw new Exception("AttributeType can not be empty.");
            }

            var lowerAttributeTypeString = attributeTypeString.ToLower();
            if (DefaultConfiguration.StringAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof (string);
            }
            else if (DefaultConfiguration.IntAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof(int);
            }
            else if (DefaultConfiguration.DecimalAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof(decimal);
            }
            else if (DefaultConfiguration.OptionSetAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof(OptionSet);
            }
            else if (DefaultConfiguration.GlobalOptionSetAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof(GlobalOptionSet);
            }
            else if (DefaultConfiguration.BoolAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof(bool);
            }
            else if (DefaultConfiguration.MoneyAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof(Money);
            }
            else if (DefaultConfiguration.DateTimeAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof(DateTime);
            }
            else if (DefaultConfiguration.LookupAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof(Lookup);
            }
            else if (DefaultConfiguration.MultilineAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof(Multiline);
            }
            else
            {
                throw new Exception("AttributeType is not in defined list.");
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
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