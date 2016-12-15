using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk.Metadata;

namespace EntityCreator.Models
{
    public class TemplateBase
    {
        public string DisplayName { get; set; }
        public string DisplayNameShort { get; set; }
        public string Description { get; set; }
        public string LogicalName { get; set; }
        public string OtherDisplayName { get; set; }
        public string OtherDescription { get; set; }
    }

    public class EntityTemplate : TemplateBase
    {
        public string DisplayNamePlural { get; set; }
        public List<AttributeTemplate> AttributeList { get; set; }
        public List<WebResource> WebResource { get; set; }
        public List<Exception> Warnings { get; set; }
        public List<Exception> Errors { get; set; }
        public bool WillCreateEntity { get; set; }
    }

    public class AttributeTemplate : TemplateBase
    {
        public bool IsRequired { get; set; }
        public int MaxLength { get; set; }
        public int MinLength { get; set; }
        public Type AttributeType { get; set; }
        public StringFormatName StringFormatName { get; set; }
        public int Precision { get; set; }
        public DateTimeFormat? DateTimeFormat { get; set; }
        public IntegerFormat? IntegerFormat { get; set; }
        public bool? BooleanDefaultValue { get; set; }
        public string LookupEntityLogicalName { get; set; }
        public string GlobalOptionSetListLogicalName { get; set; }
        public List<OptionSetTemplate> OptionSetList { get; set; }
    }

    public class OptionSetTemplate
    {
        public string Description { get; set; }
        public string Label { get; set; }
        public int Value { get; set; }
    }
}