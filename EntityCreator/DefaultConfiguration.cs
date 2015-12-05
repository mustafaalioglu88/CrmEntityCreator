using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk.Metadata;

namespace EntityCreator
{
    public static class DefaultConfiguration
    {
        // IsMultipleExecuteRequest overwrites IsMultiThreadSupport configuration
        public static bool IsMultipleExecuteRequest = true;
        public static bool IsMultiThreadSupport = false;
        // Zero to close
        public const int TimeoutBeforeNextExecuteInSeconds = 0;
        public const int DefaultLanguageCode = 1055;
        public static bool ThrowExceptionOnNegligibleErrors = false;
        public const int TimeoutBeforeExceptionInMinutes = 60;
        public const OwnershipTypes DefaultOwnershipType = OwnershipTypes.UserOwned;
        public const int DefaultStringMaxLength = 100;
        public const DateTimeFormat DefaultDateTimeFormat = DateTimeFormat.DateOnly;
        public const int DefaultMoneyPrecisionSource = 2;
        public const IntegerFormat DefaultIntegerFormat = IntegerFormat.None;
        public const int DefaultDecimalPrecision = 2;
        public const int DefaultIntMinValue = Int32.MinValue;
        public const int DefaultIntMaxValue = Int32.MaxValue;
        public const decimal DefaultDecimalMinValue = Decimal.MinValue;
        public const decimal DefaultDecimalMaxValue = Decimal.MaxValue;
        public const bool DefaultBooleanDefaultValue = false;
        public const string DefaultPrimaryAttribute = "tse_name";
        public const string DefaultPrimaryAttributeDisplayName = "Name";
        public const string DefaultPrimaryAttributeDescription = "Primary field for entity.";
        public static readonly StringFormatName DefaultStringFormatName = StringFormatName.Text;
        public static readonly List<string> StringAttributeTypeList = new List<string>() { "string" }; 
        public static readonly List<string> IntAttributeTypeList = new List<string>() { "int" }; 
        public static readonly List<string> DecimalAttributeTypeList = new List<string>() { "decimal" };
        public static readonly List<string> OptionSetAttributeTypeList = new List<string>() { "optionset" };
        public static readonly List<string> GlobalOptionSetAttributeTypeList = new List<string>() { "globaloptionset" };
        public static readonly List<string> BoolAttributeTypeList = new List<string>() { "bool" };
        public static readonly List<string> MoneyAttributeTypeList = new List<string>() { "money" };
        public static readonly List<string> DateTimeAttributeTypeList = new List<string>() { "datetime" };
        public static readonly List<string> LookupAttributeTypeList = new List<string>() { "lookup" };
        public static readonly List<string> MultilineAttributeTypeList = new List<string>() { "multiline" };
        public const string YesDefaultValue = "Evet";
        public const string NoDefaultValue = "Hayır";
        public const int DefaultMemoMaxLength = 2000;
        public static readonly char[] OptionSetSplicChar = { ';' };
        public const string SolutionUniqueName = "";
        public const string WebResourceTemplate = "<html><head><meta><meta><meta><meta></head><body style='word-wrap: break-word;'><p style='font-family: Segoe UI,Tahoma,Arial; font-size: 12px;'>{0}</p></body></html>";
        public const int AttributeDescriptionMaxLength = 250;
        public const int AttributeDisplayNameMaxLength = 100;
        public const int AttributeLogicalMaxLength = 100;
    }
}