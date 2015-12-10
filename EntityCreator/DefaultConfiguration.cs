using System.Collections.Generic;
using EntityCreator.Models;
using Microsoft.Xrm.Sdk.Metadata;

namespace EntityCreator
{
    public static class DefaultConfiguration
    {
        public static readonly List<string> MessagesOfTheDay = new List<string>
        {
            "What a lovely day",
            "You look very nice today",
            "You deserve everything",
            "Everything is for you",
            "You can do it",
            "Just do it",
            "Make your wishes come true"
        };

        #region ExecutionConfigurationParameters

        public const int DefaultLanguageCode = 1055;
        public const int ExecuteMultipleSize = 10;
        public const int TimeoutExceptionInMinutes = 60;
        public const int AttributeDisplayNameMaxLength = 100;
        public const string ConnectionFormat = "Url={0}; Domain={1}; Username={2}; Password={3};";

        // Not necessary to add
        public const string SolutionUniqueName = "";

        #endregion

        #region DefaultSettings

        public const int DefaultStringMaxLength = 100;
        public const int DefaultMemoMaxLength = 2000;
        public const int DefaultMoneyPrecisionSource = 2;
        public const int DefaultDecimalPrecision = 2;
        public const int DefaultDoublePrecision = 2;
        public const int DefaultIntMinValue = int.MinValue;
        public const int DefaultIntMaxValue = int.MaxValue;
        public const decimal DefaultDecimalMinValue = decimal.MinValue;
        public const decimal DefaultDecimalMaxValue = decimal.MaxValue;
        public const double DefaultDoubleMinValue = double.MinValue;
        public const double DefaultDoubleMaxValue = double.MaxValue;
        public const bool DefaultBooleanDefaultValue = false;
        public const OwnershipTypes DefaultOwnershipType = OwnershipTypes.UserOwned;
        public const DateTimeFormat DefaultDateTimeFormat = DateTimeFormat.DateOnly;
        public const IntegerFormat DefaultIntegerFormat = IntegerFormat.None;
        public static readonly StringFormatName DefaultStringFormatName = StringFormatName.Text;

        public const string YesDefaultValue = "Evet";
        public const string NoDefaultValue = "Hayır";

        public const string DefaultPrimaryAttribute = "tse_name";
        public const string DefaultPrimaryAttributeDisplayName = "Name";
        public const string DefaultPrimaryAttributeDescription = "Primary field for entity.";

        #endregion

        #region AttributeTypeStringLists

        public static readonly List<string> StringAttributeTypeList = new List<string> {"string"};
        public static readonly List<string> IntAttributeTypeList = new List<string> {"int"};
        public static readonly List<string> DecimalAttributeTypeList = new List<string> {"decimal"};
        public static readonly List<string> OptionSetAttributeTypeList = new List<string> {"optionset"};
        public static readonly List<string> GlobalOptionSetAttributeTypeList = new List<string> {"globaloptionset"};
        public static readonly List<string> BoolAttributeTypeList = new List<string> {"bool"};
        public static readonly List<string> MoneyAttributeTypeList = new List<string> {"money"};
        public static readonly List<string> DateTimeAttributeTypeList = new List<string> {"datetime"};
        public static readonly List<string> LookupAttributeTypeList = new List<string> {"lookup"};
        public static readonly List<string> MultilineAttributeTypeList = new List<string> {"multiline"};
        public static readonly List<string> FloatAttributeTypeList = new List<string> {"float", "double"};

        public static readonly char[] OptionSetSplicChar = {';'};

        public static readonly Dictionary<string, int> WebresourceTypes = new Dictionary<string, int>
        {
            {"html", (int) Enums.WebResourceTypes.Html},
            {"css", (int) Enums.WebResourceTypes.Css},
            {"jscript", (int) Enums.WebResourceTypes.JScript},
            {"xml", (int) Enums.WebResourceTypes.Xml}
        };

        public const string WebResourceHtmlTemplate = 
            "<html>" +
            "<head><meta><meta><meta><meta></head>" +
            "<body style='word-wrap: break-word;'>" +
            "<p style='word-wrap: break-word;color:#444444;font-size:12px;font-family:Segoe\000020UI,Tahoma,Arial;overflow:hidden'>" +
            "{0}" +
            "</p>" +
            "</body>" +
            "</html>";

        #endregion
    }
}