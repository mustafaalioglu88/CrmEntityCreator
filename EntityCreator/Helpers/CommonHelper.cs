using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EntityCreator.Models;
using Microsoft.Xrm.Sdk;

namespace EntityCreator.Helpers
{
    public static class CommonHelper
    {
        public static string EncodeTo64(string toEncode)
        {
            var toEncodeAsBytes = Encoding.UTF8.GetBytes(toEncode);
            var returnValue = Convert.ToBase64String(toEncodeAsBytes);
            return returnValue;
        }

        public static int GetWebResourceType(string type, string excelFile, List<Exception> errorList,
            List<Exception> warningList)
        {
            if (string.IsNullOrWhiteSpace(type))
            {
                errorList.Add(new Exception(string.Format("Webresource type can not be empty. File: {0}", excelFile)));
                return -1;
            }

            var lowerstringType = type.ToLower();
            return DefaultConfiguration.WebresourceTypes.First(x => x.Key == lowerstringType).Value;
        }

        public static Type GetAttributeType(string attributeTypeString, string excelFile, List<Exception> errorList,
            List<Exception> warningList)
        {
            if (string.IsNullOrWhiteSpace(attributeTypeString))
            {
                errorList.Add(new Exception(string.Format("AttributeType can not be empty. File: {0}", excelFile)));
                return typeof (Exception);
            }

            var lowerAttributeTypeString = attributeTypeString.ToLower();
            if (DefaultConfiguration.StringAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof (string);
            }
            if (DefaultConfiguration.IntAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof (int);
            }
            if (DefaultConfiguration.DecimalAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof (decimal);
            }
            if (DefaultConfiguration.OptionSetAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof (OptionSet);
            }
            if (DefaultConfiguration.GlobalOptionSetAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof (GlobalOptionSet);
            }
            if (DefaultConfiguration.BoolAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof (bool);
            }
            if (DefaultConfiguration.MoneyAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof (Money);
            }
            if (DefaultConfiguration.DateTimeAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof (DateTime);
            }
            if (DefaultConfiguration.LookupAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof (Lookup);
            }
            if (DefaultConfiguration.MultilineAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof (Multiline);
            }
            if (DefaultConfiguration.FloatAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof(float);
            }
            if (DefaultConfiguration.PrimaryAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof(Primary);
            }
            if (DefaultConfiguration.NNRelationAttributeTypeList.Contains(lowerAttributeTypeString))
            {
                return typeof(NNRelation);
            }
            errorList.Add(new Exception(string.Format("AttributeType is not in defined list. File: {0} Type: {1}", excelFile, lowerAttributeTypeString)));
            return typeof (Exception);
        }
    }
}