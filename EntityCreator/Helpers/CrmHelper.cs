﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using EntityCreator.Models;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace EntityCreator.Helpers
{
    public class CrmHelper: IDisposable
    {
        private List<CreateAttributeRequest> createAttributeRequestList;
        private List<string> createdGlobalOptionSetList;
        private List<CreateRequest> createWebResourcesRequestList;
        private readonly CrmConnection crmConnection;
        private List<Exception> errorList;
        private List<Exception> warningList;
        private OrganizationService sharedOrganizationService;

        public CrmHelper(string url, string domain, string username, string password)
        {
            crmConnection =
                CrmConnection.Parse(string.Format(DefaultConfiguration.ConnectionFormat, url, domain, username, password));
            crmConnection.Timeout = new TimeSpan(0, 0, DefaultConfiguration.TimeoutExceptionInMinutes, 0);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private OrganizationService GetSharedOrganizationService()
        {
            return sharedOrganizationService ?? (sharedOrganizationService = new OrganizationService(crmConnection));
        }

        private void CreateAttribute(string entityLogicalName, AttributeTemplate attributeTemplate)
        {
            if (attributeTemplate.AttributeType == typeof (Lookup))
            {
                CreateLookupAttribute(entityLogicalName, attributeTemplate);
                return;
            }

            if (attributeTemplate.AttributeType == typeof(NNRelation))
            {
                CreateNNRelation(entityLogicalName, attributeTemplate);
                return;
            }

            if (attributeTemplate.AttributeType == typeof (GlobalOptionSet) &&
                !createdGlobalOptionSetList.Contains(attributeTemplate.GlobalOptionSetListLogicalName))
            {
                CreateGlobalOptionSetAttribute(attributeTemplate);
            }

            CreateRequest createWebRequest = null;
            if (attributeTemplate.DisplayName.Length > DefaultConfiguration.AttributeDisplayNameMaxLength)
            {
                createWebRequest = GetCreateWebResourceRequest(entityLogicalName, attributeTemplate);
            }
            var createAttributeRequest = GetCreateAttributeRequest(entityLogicalName, attributeTemplate);
            if (createAttributeRequest != null)
            {
                createAttributeRequestList.Add(createAttributeRequest);
                if (createWebRequest != null)
                {
                    createWebResourcesRequestList.Add(createWebRequest);
                }
            }
        }

        private void ExecuteMultipleWebresourceRequests(ExecuteMultipleRequest request)
        {
            try
            {
                var organizationService = GetSharedOrganizationService();
                ExecuteExecuteMultipleRequest(organizationService, request);
            }
            catch (Exception ex)
            {
                var exception = new Exception("Error occured while creating webresources", ex);
                errorList.Add(exception);
            }
        }

        private void ExecuteMultipleOperation(string exceptionMessage, Action<string> setMessageOfTheDayLabel)
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
                    for (var i = 0; i < createAttributeRequestList.Count; i++)
                    {
                        executeMultipleRequest.Requests.Add(createAttributeRequestList[i]);
                        if (i%DefaultConfiguration.ExecuteMultipleSize == 0 || createAttributeRequestList.Count-i<10)
                        {
                            ExecuteExecuteMultipleRequest(organizationService, executeMultipleRequest);
                            //setMessageOfTheDayLabel(executeMultipleRequest.Requests.Count);
                            executeMultipleRequest = new ExecuteMultipleRequest
                            {
                                Requests = new OrganizationRequestCollection()
                            };
                        }
                    }
                }

                if (createWebResourcesRequestList.Any())
                {
                    var executeMultipleRequest = new ExecuteMultipleRequest
                    {
                        Requests = new OrganizationRequestCollection()
                    };
                    for (var i = 0; i < createWebResourcesRequestList.Count; i++)
                    {
                        executeMultipleRequest.Requests.Add(createWebResourcesRequestList[i]);
                        if (i % DefaultConfiguration.ExecuteMultipleSize == 0 || createAttributeRequestList.Count - i < 10)
                        {
                            ExecuteExecuteMultipleRequest(organizationService, executeMultipleRequest);
                            //setMessageOfTheDayLabel(executeMultipleRequest.Requests.Count);
                            executeMultipleRequest = new ExecuteMultipleRequest
                            {
                                Requests = new OrganizationRequestCollection()
                            };
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var exception = new Exception(exceptionMessage, ex);
                errorList.Add(exception);
            }
        }

        private void ExecuteExecuteMultipleRequest(OrganizationService organizationService,
            ExecuteMultipleRequest executeMultipleRequest)
        {
            executeMultipleRequest.Settings = new ExecuteMultipleSettings
            {
                ContinueOnError = true,
                ReturnResponses = true
            };
            var organizationResponse = (ExecuteMultipleResponse) organizationService.Execute(executeMultipleRequest);
            if (organizationResponse == null || organizationResponse.Responses == null)
            {
                errorList.Add(new Exception("Failed to get response for ExecuteMultipleOperation."));
                return;
            }

            foreach (var responseItem in organizationResponse.Responses)
            {
                if (responseItem.Fault != null)
                {
                    var index = responseItem.RequestIndex;
                    var schemaName =
                        ((CreateAttributeRequest) executeMultipleRequest.Requests[index]).Attribute.SchemaName;
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

        private void ExecuteOperation(OrganizationService organizationService,
            OrganizationRequest createAttributeRequest, string exceptionMessage)
        {
            try
            {
                organizationService.Execute(createAttributeRequest);
            }
            catch (FaultException<OrganizationServiceFault> ex)
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
                }
            }
            catch (Exception ex)
            {
                var exception = new Exception(exceptionMessage, ex);
                errorList.Add(exception);
            }
        }

        private CreateAttributeRequest GetCreateAttributeRequest(string entityLogicalName,
            AttributeTemplate attributeTemplate)
        {
            if (attributeTemplate.AttributeType == typeof(Primary))
            {
                return null;
            }

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
            else if (attributeTemplate.AttributeType == typeof (GlobalOptionSet))
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
            else if (attributeTemplate.AttributeType == typeof (Multiline))
            {
                createAttributeRequest.Attribute = CreateMultilineAttributeMetadata(attributeTemplate);
            }
            else if (attributeTemplate.AttributeType == typeof(float))
            {
                createAttributeRequest.Attribute = CreateFloatAttributeMetadata(attributeTemplate);
            }
            else
            {
                var exception =
                    new Exception(string.Format("Given attribute type is not supported. Type: {0}",
                        attributeTemplate.AttributeType));
                errorList.Add(exception);
                return null;
            }

            createAttributeRequest.Attribute.SchemaName = attributeTemplate.LogicalName;
            createAttributeRequest.Attribute.RequiredLevel =
                new AttributeRequiredLevelManagedProperty(attributeTemplate.IsRequired
                    ? AttributeRequiredLevel.SystemRequired
                    : AttributeRequiredLevel.None);
            createAttributeRequest.Attribute.DisplayName = GetLabelWithLocalized(attributeTemplate.DisplayNameShort);
            createAttributeRequest.Attribute.Description = GetLabelWithLocalized(attributeTemplate.Description);
            if(!string.IsNullOrWhiteSpace(attributeTemplate.OtherDisplayName))
            {
                var otherDisplayLabel = new LocalizedLabel(attributeTemplate.OtherDisplayName, DefaultConfiguration.OtherLanguageCode);
                createAttributeRequest.Attribute.DisplayName.LocalizedLabels.Add(otherDisplayLabel);
            }

            if (!string.IsNullOrWhiteSpace(attributeTemplate.OtherDescription))
            {
                var otherDescriptionLabel = new LocalizedLabel(attributeTemplate.OtherDescription, DefaultConfiguration.OtherLanguageCode);
                createAttributeRequest.Attribute.Description.LocalizedLabels.Add(otherDescriptionLabel);
            }

            return createAttributeRequest;
        }
        
        private CreateRequest GetCreateWebResourceRequest(string entityLogicalName, AttributeTemplate attributeTemplate)
        {
            var contents =
                CommonHelper.EncodeTo64(string.Format(DefaultConfiguration.WebResourceHtmlTemplate,
                    attributeTemplate.DisplayName));
            var webResource = new WebResource
            {
                Content = contents,
                DisplayName = attributeTemplate.DisplayNameShort,
                Description = attributeTemplate.Description,
                Name = entityLogicalName + "_" + attributeTemplate.LogicalName,
                LogicalName = WebResource.EntityLogicalName,
                WebResourceType = new OptionSetValue((int) Enums.WebResourceTypes.Html)
            };
            var createRequest = new CreateRequest
            {
                Target = webResource
            };

            if (!string.IsNullOrWhiteSpace(DefaultConfiguration.SolutionUniqueName))
            {
                createRequest.Parameters.Add("SolutionUniqueName", DefaultConfiguration.SolutionUniqueName);
            }

            return createRequest;
        }

        private CreateRequest GetCreateWebResourceRequest(WebResource resource)
        {
            var createRequest = new CreateRequest
            {
                Target = resource
            };

            if (!string.IsNullOrWhiteSpace(DefaultConfiguration.SolutionUniqueName))
            {
                createRequest.Parameters.Add("SolutionUniqueName", DefaultConfiguration.SolutionUniqueName);
            }

            return createRequest;
        }

        private DoubleAttributeMetadata CreateFloatAttributeMetadata(AttributeTemplate attributeTemplate)
        {
            var floatAttributeMetadata = new DoubleAttributeMetadata 
            {
                MinValue = -100000000000.00,
                MaxValue = 100000000000.00,
                Precision = 2
            };

            return floatAttributeMetadata;
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

        private PicklistAttributeMetadata CreateGlobalOptionSetAttributeMetadata(AttributeTemplate attributeTemplate)
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
                    DisplayName = GetLabelWithLocalized(attributeTemplate.DisplayNameShort),
                    Description = GetLabelWithLocalized(attributeTemplate.Description),
                    IsGlobal = true,
                    OptionSetType = OptionSetType.Picklist
                }
            };

            if (!string.IsNullOrWhiteSpace(attributeTemplate.OtherDisplayName))
            {
                var otherDisplayLabel = new LocalizedLabel(attributeTemplate.OtherDisplayName, DefaultConfiguration.OtherLanguageCode);
                createOptionSetRequest.OptionSet.DisplayName.LocalizedLabels.Add(otherDisplayLabel);
            }

            if (!string.IsNullOrWhiteSpace(attributeTemplate.OtherDescription))
            {
                var otherDescriptionLabel = new LocalizedLabel(attributeTemplate.OtherDescription, DefaultConfiguration.OtherLanguageCode);
                createOptionSetRequest.OptionSet.Description.LocalizedLabels.Add(otherDescriptionLabel);
            }

            ExecuteOperation(GetSharedOrganizationService(), createOptionSetRequest,
                string.Format("An error occured while creating the attribute: {0}",
                    attributeTemplate.LogicalName));

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
                MinValue = -100000000000.00M,
                MaxValue = 100000000000.00M,
                Precision = 2
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
            var schemaName = attributeTemplate.LogicalName + "_" +
                        attributeTemplate.LookupEntityLogicalName + "_" +
                        entityLogicalName;
            if(schemaName.Length > 100)
            {
                schemaName = schemaName.Substring(default(int), 100);
            }

            var createOneToManyRequest = new CreateOneToManyRequest
            {
                Lookup = new LookupAttributeMetadata
                {
                    Description = GetLabelWithLocalized(attributeTemplate.Description),
                    DisplayName = GetLabelWithLocalized(attributeTemplate.DisplayNameShort),
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
                    SchemaName = schemaName
                }
            };

            if (!string.IsNullOrWhiteSpace(attributeTemplate.OtherDisplayName))
            {
                var otherDisplayLabel = new LocalizedLabel(attributeTemplate.OtherDisplayName, DefaultConfiguration.OtherLanguageCode);
                createOneToManyRequest.Lookup.DisplayName.LocalizedLabels.Add(otherDisplayLabel);
            }

            if (!string.IsNullOrWhiteSpace(attributeTemplate.OtherDescription))
            {
                var otherDescriptionLabel = new LocalizedLabel(attributeTemplate.OtherDescription, DefaultConfiguration.OtherLanguageCode);
                createOneToManyRequest.Lookup.Description.LocalizedLabels.Add(otherDescriptionLabel);
            }

            CreateRequest createWebRequest = null;
            if (attributeTemplate.DisplayName.Length > DefaultConfiguration.AttributeDisplayNameMaxLength)
            {
                createWebRequest = GetCreateWebResourceRequest(entityLogicalName, attributeTemplate);
            }

            ExecuteOperation(GetSharedOrganizationService(), createOneToManyRequest,
                string.Format("An error occured while creating the attribute: {0}",
                    attributeTemplate.LogicalName));

            if (createWebRequest != null)
            {
                ExecuteOperation(GetSharedOrganizationService(), createWebRequest,
                    string.Format("An error occured while creating the web resource for attribute: {0}",
                        attributeTemplate.LogicalName));
            }
        }

        private void CreateNNRelation(string entityLogicalName, AttributeTemplate attributeTemplate)
        {
            string entity1Logical, entity2Logical;
            var displayNames = attributeTemplate.DisplayName.Split(';');
            var entity1Display = displayNames.First();
            var entity2Display = displayNames.Last();
            if (attributeTemplate.Description == "1")
            {
                entity1Logical = attributeTemplate.LogicalName;
                entity2Logical = attributeTemplate.LookupEntityLogicalName;
            }
            else
            {
                entity1Logical = attributeTemplate.LookupEntityLogicalName;
                entity2Logical = attributeTemplate.LogicalName;
            }

            AssociatedMenuBehavior? entity1Behavior = entity1Display == entity1Logical ? AssociatedMenuBehavior.DoNotDisplay : AssociatedMenuBehavior.UseLabel;
            AssociatedMenuBehavior? entity2Behavior = entity2Display == entity2Logical ? AssociatedMenuBehavior.DoNotDisplay : AssociatedMenuBehavior.UseLabel;

            var createManyToManyRelationshipRequest =
                new CreateManyToManyRequest
                {
                    IntersectEntitySchemaName = attributeTemplate.LogicalName,
                    ManyToManyRelationship = new ManyToManyRelationshipMetadata
                    {
                        SchemaName = attributeTemplate.LogicalName,
                        Entity1LogicalName = entity1Logical,
                        Entity1AssociatedMenuConfiguration =
                        new AssociatedMenuConfiguration
                        {
                            Behavior = entity1Behavior,
                            Group = AssociatedMenuGroup.Details,
                            Label = new Label(entity1Display, 1033),
                            Order = 10000
                        },
                        Entity2LogicalName = entity2Logical,
                        Entity2AssociatedMenuConfiguration =
                        new AssociatedMenuConfiguration
                        {
                            Behavior = entity2Behavior,
                            Group = AssociatedMenuGroup.Details,
                            Label = new Label(entity2Display, 1033),
                            Order = 10000
                        }
                    }
                };

            ExecuteOperation(GetSharedOrganizationService(), createManyToManyRelationshipRequest,
                string.Format("An error occured while creating the NN: {0}",
                    attributeTemplate.LogicalName));

        }

        public void CreateEntity(string excelFile, EntityTemplate entityTemplate, Action<string> setMessageOfTheDayLabel)
        {
            errorList = new List<Exception>();
            warningList = new List<Exception>();
            createAttributeRequestList = new List<CreateAttributeRequest>();
            createdGlobalOptionSetList = new List<string>();
            createWebResourcesRequestList = new List<CreateRequest>();

            if (entityTemplate == null)
            {
                errorList.Add(
                    new Exception(string.Format("CreateEntity method EntityTemplate reference is null. File: {0}",
                        excelFile)));
                entityTemplate=new EntityTemplate(){Errors = new List<Exception>()};
                entityTemplate.Errors.AddRange(errorList);
                return;
            }
            if (entityTemplate.WillCreateEntity)
            {
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
                createrequest.PrimaryAttribute = GetPrimaryAttribute(entityTemplate);
                createrequest.Entity = entityMetadata;
                ExecuteOperation(GetSharedOrganizationService(), createrequest,
                    string.Format("An error occured while creating the entity: {0}", entityTemplate.LogicalName));
                //setMessageOfTheDayLabel("1");
                if (entityTemplate.AttributeList == null)
                {
                    entityTemplate.Warnings.AddRange(warningList);
                    entityTemplate.Errors.AddRange(errorList);
                }


                foreach (var attributeTemplate in entityTemplate.AttributeList)
                {
                    CreateAttribute(entityTemplate.LogicalName, attributeTemplate);
                }
            }
            
            var requests = new ExecuteMultipleRequest
            {
                Requests = new OrganizationRequestCollection()
            };
            if (entityTemplate.WebResource != null)
            {
                for (var i = 0; i < entityTemplate.WebResource.Count; i++)
                {
                    var webresourceRequest = GetCreateWebResourceRequest(entityTemplate.WebResource[i]);
                    requests.Requests.Add(webresourceRequest);
                    if (i % DefaultConfiguration.ExecuteMultipleSize == 0 || createAttributeRequestList.Count - i < 10)
                    {
                        ExecuteMultipleWebresourceRequests(requests);
                        //setMessageOfTheDayLabel(requests.Requests.Count);
                        requests = new ExecuteMultipleRequest
                        {
                            Requests = new OrganizationRequestCollection()
                        };
                    }
                }
            }

            ExecuteMultipleOperation("Execute multiple error occured.", setMessageOfTheDayLabel);
            entityTemplate.Warnings.AddRange(warningList);
            entityTemplate.Errors.AddRange(errorList);
        }

        private StringAttributeMetadata GetPrimaryAttribute(EntityTemplate entityTemplate)
        {
            var primaryAttribute = (from attribute in entityTemplate.AttributeList where attribute.AttributeType == typeof(Primary) select attribute).FirstOrDefault();

            var stringAttributeMetadata = new StringAttributeMetadata
            {
                SchemaName = primaryAttribute != null ? primaryAttribute.LogicalName : DefaultConfiguration.DefaultPrimaryAttribute,
                RequiredLevel = new AttributeRequiredLevelManagedProperty(
                    primaryAttribute != null && primaryAttribute.IsRequired ? AttributeRequiredLevel.SystemRequired : AttributeRequiredLevel.None),
                DisplayName = GetLabelWithLocalized(primaryAttribute != null ? primaryAttribute.DisplayNameShort : DefaultConfiguration.DefaultPrimaryAttributeDisplayName),
                Description = GetLabelWithLocalized(primaryAttribute != null ? primaryAttribute.Description : DefaultConfiguration.DefaultPrimaryAttributeDescription),
                MaxLength = primaryAttribute != null ? primaryAttribute.MaxLength : DefaultConfiguration.DefaultStringMaxLength,
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