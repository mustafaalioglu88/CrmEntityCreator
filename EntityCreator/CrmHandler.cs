using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace EntityCreator
{
    public class CrmHandler : IDisposable
    {
        private const string connectionFormat = "Url={0}; Domain={1}; Username={2}; Password={3};";
        private readonly CrmConnection crmConnection;
        private readonly List<Exception> errorList = new List<Exception>();
        private readonly List<Exception> warningList = new List<Exception>();
        private readonly List<string> createdGlobalOptionSetList = new List<string>();
        private readonly List<CreateAttributeRequest> createAttributeRequestList = new List<CreateAttributeRequest>();
        private OrganizationService sharedOrganizationService;

        public CrmHandler(string url, string domain, string username, string password)
        {
            crmConnection = CrmConnection.Parse(string.Format(connectionFormat, url, domain, username, password));
            crmConnection.Timeout = new TimeSpan(0, 0, DefaultConfiguration.TimeoutBeforeExceptionInMinutes, 0);
        }

        private OrganizationService GetSharedOrganizationService()
        {
            return sharedOrganizationService ?? (sharedOrganizationService = new OrganizationService(crmConnection));
        }

        private void CreateAttribute(string entityLogicalName, AttributeTemplate attributeTemplate)
        {
            if (attributeTemplate.AttributeType == typeof(Lookup))
            {
                CreateLookupAttribute(entityLogicalName, attributeTemplate);
                return;
            }

            if (attributeTemplate.AttributeType == typeof(GlobalOptionSet) && 
                !createdGlobalOptionSetList.Contains(attributeTemplate.GlobalOptionSetListLogicalName))
            {
                CreateGlobalOptionSetAttribute(attributeTemplate);
            }

            var createAttributeRequest = GetCreateAttributeRequest(entityLogicalName, attributeTemplate);
            if (createAttributeRequest != null)
            {
                if (DefaultConfiguration.IsMultipleExecuteRequest)
                {
                    createAttributeRequestList.Add(createAttributeRequest);
                }
                else if (DefaultConfiguration.IsMultiThreadSupport)
                {
                    using (var organizationService = new OrganizationService(crmConnection))
                    {
                        ExecuteOperation(organizationService, createAttributeRequest,
                            string.Format("An error occured while creating the attribute: {0}",
                                attributeTemplate.LogicalName));
                    }
                }
                else
                {
                    ExecuteOperation(GetSharedOrganizationService(), createAttributeRequest,
                            string.Format("An error occured while creating the attribute: {0}",
                                attributeTemplate.LogicalName));
                }
            }
        }

        private void ExecuteMultipleOperation(string exceptionMessage)
        {
            var organizationService = GetSharedOrganizationService();
            try
            {
                if (createAttributeRequestList.Any())
                {
                    var executeMultipleRequest = new ExecuteMultipleRequest
                    {
                        Requests = new OrganizationRequestCollection()
                    };
                    executeMultipleRequest.Requests.AddRange(createAttributeRequestList);
                    executeMultipleRequest.Settings = new ExecuteMultipleSettings
                    {
                        ContinueOnError = true,
                        ReturnResponses = true
                    };
                    var organizationResponse = (ExecuteMultipleResponse) organizationService.Execute(executeMultipleRequest);
                    if (organizationResponse == null || organizationResponse.Responses == null)
                    {
                        throw new Exception("Failed to get response for ExecuteMultipleOperation.");
                    }

                    foreach (var responseItem in organizationResponse.Responses)
                    {
                        if (responseItem.Fault != null)
                        {
                            var index = responseItem.RequestIndex;
                            var schemaName = ((CreateAttributeRequest)executeMultipleRequest.Requests[index]).Attribute.SchemaName;
                            var errorMessage = string.Format(
                                    "An error occured while creating the attribute: {0}. Detailed error:\n{1}",
                                    schemaName, responseItem.Fault.ToErrorString());
                            if (responseItem.Fault.ErrorCode == -2147192813)
                            {
                                warningList.Add(new Exception(errorMessage));
                            }
                            else
                            {
                                errorList.Add(new Exception(errorMessage));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var exception = new Exception(exceptionMessage, ex);
                errorList.Add(exception);
                if (DefaultConfiguration.ThrowExceptionOnNegligibleErrors)
                {
                    throw exception;
                }
            }
        }
        private void ExecuteOperation(OrganizationService organizationService, OrganizationRequest createAttributeRequest, string exceptionMessage)
        {
            try
            {
                organizationService.Execute(createAttributeRequest);
            }
            catch (System.ServiceModel.FaultException<OrganizationServiceFault> ex)
            {
                var exception = new Exception(exceptionMessage, ex);
                if (ex.Detail.ErrorCode == -2147192821 ||
                    ex.Detail.ErrorCode == -2147204283 || 
                    ex.Detail.ErrorCode == -2147192813)
                {
                    warningList.Add(exception);
                }
                else
                {
                    errorList.Add(exception);
                    if (DefaultConfiguration.ThrowExceptionOnNegligibleErrors)
                    {
                        throw exception;
                    }
                }
            }
            catch(Exception ex)
            {
                var exception = new Exception(exceptionMessage, ex);
                errorList.Add(exception);
                if (DefaultConfiguration.ThrowExceptionOnNegligibleErrors)
                {
                    throw exception;
                }
            }

            if (DefaultConfiguration.TimeoutBeforeNextExecuteInSeconds > default(int))
            {
                Thread.Sleep(new TimeSpan(0, 0, DefaultConfiguration.TimeoutBeforeNextExecuteInSeconds));
            }
        }

        private CreateAttributeRequest GetCreateAttributeRequest(string entityLogicalName,
            AttributeTemplate attributeTemplate)
        {
            var createAttributeRequest = new CreateAttributeRequest {EntityName = entityLogicalName};
            if (attributeTemplate.AttributeType == typeof (string))
            {
                createAttributeRequest.Attribute = CreateStringAttributeMetadata(attributeTemplate);
            }
            else if (attributeTemplate.AttributeType == typeof (int))
            {
                createAttributeRequest.Attribute = CreateIntAttributeMetadata(attributeTemplate);
            }
            else if (attributeTemplate.AttributeType == typeof (decimal))
            {
                createAttributeRequest.Attribute = CreateDecimalAttributeMetadata(attributeTemplate);
            }
            else if (attributeTemplate.AttributeType == typeof (OptionSet))
            {
                createAttributeRequest.Attribute = CreateOptionSetAttributeMetadata(attributeTemplate);
            }
            else if (attributeTemplate.AttributeType == typeof(GlobalOptionSet))
            {
                createAttributeRequest.Attribute = CreateGlobalOptionSetAttributeMetadata(attributeTemplate);
            }
            else if (attributeTemplate.AttributeType == typeof (bool))
            {
                createAttributeRequest.Attribute = CreateBoolAttributeMetadata(attributeTemplate);
            }
            else if (attributeTemplate.AttributeType == typeof (Money))
            {
                createAttributeRequest.Attribute = CreateMoneyAttributeMetadata(attributeTemplate);
            }
            else if (attributeTemplate.AttributeType == typeof (DateTime))
            {
                createAttributeRequest.Attribute = CreateDateTimeAttributeMetadata(attributeTemplate);
            }
            else if (attributeTemplate.AttributeType == typeof(Multiline))
            {
                createAttributeRequest.Attribute = CreateMultilineAttributeMetadata(attributeTemplate);
            }
            else
            {
                var exception = new Exception(string.Format("Given attribute type is not supported. Type: {0}",
                    attributeTemplate.AttributeType));
                errorList.Add(exception);
                if (DefaultConfiguration.ThrowExceptionOnNegligibleErrors)
                {
                    throw exception;
                }
                else
                {
                    return null;
                }
            }

            createAttributeRequest.Attribute.SchemaName = attributeTemplate.LogicalName;
            createAttributeRequest.Attribute.RequiredLevel =
                new AttributeRequiredLevelManagedProperty(attributeTemplate.IsRequired
                    ? AttributeRequiredLevel.SystemRequired
                    : AttributeRequiredLevel.None);
            createAttributeRequest.Attribute.DisplayName = GetLabelWithLocalized(attributeTemplate.DisplayName);
            createAttributeRequest.Attribute.Description = GetLabelWithLocalized(attributeTemplate.Description);
            return createAttributeRequest;
        }

        private MemoAttributeMetadata CreateMultilineAttributeMetadata(AttributeTemplate attributeTemplate)
        {
            var stringAttributeMetadata = new MemoAttributeMetadata
            {
                MaxLength = attributeTemplate.MaxLength == default(int)
                    ? DefaultConfiguration.DefaultMemoMaxLength
                    : attributeTemplate.MaxLength,
                Format = StringFormat.TextArea
            };
            return stringAttributeMetadata;
        }

        private AttributeMetadata CreateGlobalOptionSetAttributeMetadata(AttributeTemplate attributeTemplate)
        {
            var picklistAttributeMetadata = new PicklistAttributeMetadata
            {
                OptionSet = new OptionSetMetadata
                {
                    IsGlobal = true,
                    Name = attributeTemplate.GlobalOptionSetListLogicalName
                }
            };
            return picklistAttributeMetadata;
        }

        private void CreateGlobalOptionSetAttribute(AttributeTemplate attributeTemplate)
        {
            var optionMetadataCollection = GetOptionMetadataCollection(attributeTemplate);
            var createOptionSetRequest = new CreateOptionSetRequest
            {
                OptionSet = new OptionSetMetadata(optionMetadataCollection)
                {
                    Name = attributeTemplate.GlobalOptionSetListLogicalName,
                    DisplayName = GetLabelWithLocalized(attributeTemplate.DisplayName),
                    IsGlobal = true,
                    OptionSetType = OptionSetType.Picklist
                }
            };

            if (DefaultConfiguration.IsMultiThreadSupport)
            {
                using (var organizationService = new OrganizationService(crmConnection))
                {
                    ExecuteOperation(organizationService, createOptionSetRequest,
                        string.Format("An error occured while creating the attribute: {0}",
                            attributeTemplate.LogicalName));
                }
            }
            else
            {
                ExecuteOperation(GetSharedOrganizationService(), createOptionSetRequest,
                        string.Format("An error occured while creating the attribute: {0}",
                            attributeTemplate.LogicalName));
            }

            createdGlobalOptionSetList.Add(attributeTemplate.GlobalOptionSetListLogicalName);
        }

        private BooleanAttributeMetadata CreateBoolAttributeMetadata(AttributeTemplate attributeTemplate)
        {
            var yesValue = DefaultConfiguration.YesDefaultValue;
            var noValue = DefaultConfiguration.NoDefaultValue;
            if (attributeTemplate.OptionSetList != null && attributeTemplate.OptionSetList.Count == 2)
            {
                yesValue = attributeTemplate.OptionSetList[0].Label;
                noValue = attributeTemplate.OptionSetList[1].Label;
            }

            var booleanAttributeMetadata = new BooleanAttributeMetadata
            {
                OptionSet = new BooleanOptionSetMetadata(
                            new OptionMetadata(GetLabelWithLocalized(yesValue), 1),
                            new OptionMetadata(GetLabelWithLocalized(noValue), 0)
                            ),
                DefaultValue = attributeTemplate.BooleanDefaultValue == default(bool?)
                    ? DefaultConfiguration.DefaultBooleanDefaultValue
                    : attributeTemplate.BooleanDefaultValue
            };
            return booleanAttributeMetadata;
        }

        private PicklistAttributeMetadata CreateOptionSetAttributeMetadata(AttributeTemplate attributeTemplate)
        {
            var optionMetadataCollection = GetOptionMetadataCollection(attributeTemplate);
            var picklistAttributeMetadata = new PicklistAttributeMetadata
            {
                OptionSet = new OptionSetMetadata(optionMetadataCollection)
            };
            picklistAttributeMetadata.OptionSet.IsGlobal = false;
            return picklistAttributeMetadata;
        }

        private OptionMetadataCollection GetOptionMetadataCollection(AttributeTemplate attributeTemplate)
        {
            var optionMetadataCollection = new OptionMetadataCollection();
            if (attributeTemplate.OptionSetList != null)
            {
                foreach (var optionMetadataTemplate in attributeTemplate.OptionSetList)
                {
                    var optionMetadata = new OptionMetadata
                    {
                        Description = GetLabelWithLocalized(optionMetadataTemplate.Description),
                        Label = GetLabelWithLocalized(optionMetadataTemplate.Label),
                        Value = optionMetadataTemplate.Value
                    };
                    optionMetadataCollection.Add(optionMetadata);
                }
            }
            return optionMetadataCollection;
        }

        private DecimalAttributeMetadata CreateDecimalAttributeMetadata(AttributeTemplate attributeTemplate)
        {
            var decimalAttributeMetadata = new DecimalAttributeMetadata
            {
                MinValue = attributeTemplate.MinLength == default(int)
                    ? DefaultConfiguration.DefaultDecimalMinValue
                    : attributeTemplate.MinLength,
                MaxValue = attributeTemplate.MaxLength == default(int)
                    ? DefaultConfiguration.DefaultDecimalMaxValue
                    : attributeTemplate.MaxLength,
                Precision = attributeTemplate.Precision == default(int)
                    ? DefaultConfiguration.DefaultDecimalPrecision
                    : attributeTemplate.Precision
            };
            return decimalAttributeMetadata;
        }

        private IntegerAttributeMetadata CreateIntAttributeMetadata(AttributeTemplate attributeTemplate)
        {
            var integerAttributeMetadata = new IntegerAttributeMetadata
            {
                MinValue = attributeTemplate.MinLength == default(int)
                    ? DefaultConfiguration.DefaultIntMinValue
                    : attributeTemplate.MinLength,
                MaxValue = attributeTemplate.MaxLength == default(int)
                    ? DefaultConfiguration.DefaultIntMaxValue
                    : attributeTemplate.MaxLength,
                Format = attributeTemplate.IntegerFormat == default(IntegerFormat?)
                    ? DefaultConfiguration.DefaultIntegerFormat
                    : attributeTemplate.IntegerFormat
            };
            return integerAttributeMetadata;
        }

        private DateTimeAttributeMetadata CreateDateTimeAttributeMetadata(AttributeTemplate attributeTemplate)
        {
            var dateTimeAttributeMetadata = new DateTimeAttributeMetadata
            {
                Format = attributeTemplate.DateTimeFormat == default(DateTimeFormat?)
                    ? DefaultConfiguration.DefaultDateTimeFormat
                    : attributeTemplate.DateTimeFormat
            };
            return dateTimeAttributeMetadata;
        }

        private MoneyAttributeMetadata CreateMoneyAttributeMetadata(AttributeTemplate attributeTemplate)
        {
            var moneyAttributeMetadata = new MoneyAttributeMetadata
            {
                PrecisionSource = attributeTemplate.Precision == default(int)
                    ? DefaultConfiguration.DefaultMoneyPrecisionSource
                    : attributeTemplate.Precision
            };
            return moneyAttributeMetadata;
        }

        private StringAttributeMetadata CreateStringAttributeMetadata(AttributeTemplate attributeTemplate)
        {
            var stringAttributeMetadata = new StringAttributeMetadata
            {
                MaxLength = attributeTemplate.MaxLength == default(int)
                    ? DefaultConfiguration.DefaultStringMaxLength
                    : attributeTemplate.MaxLength,
                FormatName = attributeTemplate.StringFormatName == default(StringFormatName)
                    ? DefaultConfiguration.DefaultStringFormatName
                    : attributeTemplate.StringFormatName
            };
            return stringAttributeMetadata;
        }

        private void CreateLookupAttribute(string entityLogicalName, AttributeTemplate attributeTemplate)
        {
            var createOneToManyRequest = new CreateOneToManyRequest
            {
                Lookup = new LookupAttributeMetadata
                {
                    Description = GetLabelWithLocalized(attributeTemplate.Description),
                    DisplayName = GetLabelWithLocalized(attributeTemplate.DisplayName),
                    LogicalName = attributeTemplate.LogicalName,
                    SchemaName = attributeTemplate.LogicalName,
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(attributeTemplate.IsRequired
                        ? AttributeRequiredLevel.SystemRequired
                        : AttributeRequiredLevel.None)
                },
                OneToManyRelationship = new OneToManyRelationshipMetadata
                {
                    ReferencedEntity = attributeTemplate.LookupEntityLogicalName,
                    ReferencingEntity = entityLogicalName,
                    SchemaName = attributeTemplate.LookupEntityLogicalName + "_" + entityLogicalName 
                }
            };

            if (DefaultConfiguration.IsMultiThreadSupport)
            {
                using (var organizationService = new OrganizationService(crmConnection))
                {
                    ExecuteOperation(organizationService, createOneToManyRequest,
                        string.Format("An error occured while creating the attribute: {0}",
                            attributeTemplate.LogicalName));
                }
            }
            else
            {
                ExecuteOperation(GetSharedOrganizationService(), createOneToManyRequest,
                        string.Format("An error occured while creating the attribute: {0}",
                            attributeTemplate.LogicalName));
            }
        }

        public List<Exception> CreateEntity(EntityTemplate entityTemplate, out List<Exception> listOfWarning)
        {
            if (entityTemplate == null)
            {
                throw new NullReferenceException("CreateEntity method EntityTemplate reference is null.");
            }

            var createrequest = new CreateEntityRequest();
            var entityMetadata = new EntityMetadata
            {
                SchemaName = entityTemplate.LogicalName,
                DisplayName = GetLabelWithLocalized(entityTemplate.DisplayName),
                DisplayCollectionName = GetLabelWithLocalized(entityTemplate.DisplayNamePlural),
                Description = GetLabelWithLocalized(entityTemplate.Description),
                OwnershipType = DefaultConfiguration.DefaultOwnershipType,
                IsActivity = false
            };
            createrequest.PrimaryAttribute = GetPrimaryAttribute();
            createrequest.Entity = entityMetadata;

            if (DefaultConfiguration.IsMultiThreadSupport)
            {
                using (var organizationService = new OrganizationService(crmConnection))
                {
                    ExecuteOperation(organizationService, createrequest,
                        string.Format("An error occured while creating the entity: {0}", entityTemplate.LogicalName));
                }
            }
            else
            {
                ExecuteOperation(GetSharedOrganizationService(), createrequest,
                        string.Format("An error occured while creating the entity: {0}", entityTemplate.LogicalName));
            }

            if (entityTemplate.AttributeList == null)
            {
                listOfWarning = warningList;
                return errorList;
            }

            foreach (var attributeTemplate in entityTemplate.AttributeList)
            {
                CreateAttribute(entityTemplate.LogicalName, attributeTemplate);
            }

            if (DefaultConfiguration.IsMultipleExecuteRequest)
            {
                ExecuteMultipleOperation("Execute multiple error occured.");
            }

            listOfWarning = warningList;
            return errorList;
        }

        private StringAttributeMetadata GetPrimaryAttribute()
        {
            var stringAttributeMetadata = new StringAttributeMetadata
            {
                SchemaName = DefaultConfiguration.DefaultPrimaryAttribute,
                RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                DisplayName = GetLabelWithLocalized(DefaultConfiguration.DefaultPrimaryAttributeDisplayName),
                Description = GetLabelWithLocalized(DefaultConfiguration.DefaultPrimaryAttributeDescription),
                MaxLength = DefaultConfiguration.DefaultStringMaxLength,
                FormatName = DefaultConfiguration.DefaultStringFormatName
            };
            return stringAttributeMetadata;
        }

        private Label GetLabelWithLocalized(string textToBind)
        {
            if (string.IsNullOrWhiteSpace(textToBind))
            {
                textToBind = string.Empty;
            }
            var label = new Label(textToBind, DefaultConfiguration.DefaultLanguageCode);
            return label;
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
                if (sharedOrganizationService != null)
                {
                    sharedOrganizationService.Dispose();
                    sharedOrganizationService = null;
                }
            }
        }

    }
}