using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text.Json;

namespace DataversePowerAutomateHelpers
{
    public class RetrieveOptions : IPlugin
    {
        /*
        Input Parameters:
        | String | EntityLogicalName    | Required | The LogicalName of the entity that contains the attribute. |
        | String | LogicalName | Optional | The LogicalName of the attribute (or column) that contains the options. If not provided, returns all option set attributes. |


        Output Parameters:
        | String | Options | A JSON string containing the options as a key-value map|

        Response structure:
        When LogicalName is provided (single attribute):
            {
                "Options": {
                    "Red": 1,
                    "Blue": 2,
                    "Green": 3
                }
            }

        When LogicalName is not provided (all attributes):
            {
                "Options": {
                    "color": {
                        "Red": 1,
                        "Blue": 2
                    },
                    "size": {
                        "Small": 1,
                        "Large": 2
                    }
                }
            }

        */


        public void Execute(IServiceProvider serviceProvider)
        {
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("RetrieveOptions plugin execution started.");

            try
            {
                tracingService.Trace("Reading input parameters...");
                if (!context.InputParameters.Contains("EntityLogicalName"))
                {
                    throw new InvalidPluginExecutionException("EntityLogicalName input parameter is required.");
                }

                string entityLogicalName = (string)context.InputParameters["EntityLogicalName"];
                string logicalName = context.InputParameters.Contains("LogicalName") ? (string)context.InputParameters["LogicalName"] : null;
                
                tracingService.Trace($"Input - EntityLogicalName: '{entityLogicalName}', LogicalName: '{logicalName ?? "(null - retrieve all)"}''");

                object optionsData;

                if (string.IsNullOrEmpty(logicalName))
                {
                    tracingService.Trace("LogicalName is null/empty - retrieving all option set attributes for the entity.");
                    var allOptionsData = new Dictionary<string, Dictionary<string, int?>>();
                    var entityMetadata = DataversePowerAutomateHelpers.Utility.GetEntityProperties(service, entityLogicalName, "Attributes");
                    tracingService.Trace($"Retrieved entity metadata. Attribute count: {entityMetadata?.Attributes?.Length ?? 0}");
                    foreach (var attribute in entityMetadata.Attributes)
                    {
                        if (attribute is PicklistAttributeMetadata p)
                        {
                            tracingService.Trace($"Processing PicklistAttributeMetadata: {p.LogicalName}, Options count: {p.OptionSet?.Options?.Count ?? 0}");
                            var attributeOptions = new Dictionary<string, int?>();
                            foreach (var o in p.OptionSet.Options)
                            {
                                var label = o.Label?.UserLocalizedLabel?.Label ?? string.Empty;
                                attributeOptions[label] = o.Value.Value;
                            }
                            allOptionsData[p.LogicalName] = attributeOptions;
                        }
                        else if (attribute is MultiSelectPicklistAttributeMetadata m)
                        {
                            tracingService.Trace($"Processing MultiSelectPicklistAttributeMetadata: {m.LogicalName}, Options count: {m.OptionSet?.Options?.Count ?? 0}");
                            var attributeOptions = new Dictionary<string, int?>();
                            foreach (var o in m.OptionSet.Options)
                            {
                                var label = o.Label?.UserLocalizedLabel?.Label ?? string.Empty;
                                attributeOptions[label] = o.Value.Value;
                            }
                            allOptionsData[m.LogicalName] = attributeOptions;
                        }
                        else if (attribute is BooleanAttributeMetadata b)
                        {
                            tracingService.Trace($"Processing BooleanAttributeMetadata: {b.LogicalName}");
                            var attributeOptions = new Dictionary<string, int?>();
                            var trueLabel = b.OptionSet.TrueOption.Label?.UserLocalizedLabel?.Label ?? "True";
                            var falseLabel = b.OptionSet.FalseOption.Label?.UserLocalizedLabel?.Label ?? "False";
                            attributeOptions[trueLabel] = b.OptionSet.TrueOption.Value;
                            attributeOptions[falseLabel] = b.OptionSet.FalseOption.Value;
                            allOptionsData[b.LogicalName] = attributeOptions;
                        }
                        else if (attribute is StateAttributeMetadata se)
                        {
                            tracingService.Trace($"Processing StateAttributeMetadata: {se.LogicalName}, Options count: {se.OptionSet?.Options?.Count ?? 0}");
                            var attributeOptions = new Dictionary<string, int?>();
                            foreach (var o in se.OptionSet.Options)
                            {
                                var label = o.Label?.UserLocalizedLabel?.Label ?? string.Empty;
                                attributeOptions[label] = o.Value.Value;
                            }
                            allOptionsData[se.LogicalName] = attributeOptions;
                        }
                        else if (attribute is StatusAttributeMetadata su)
                        {
                            tracingService.Trace($"Processing StatusAttributeMetadata: {su.LogicalName}, Options count: {su.OptionSet?.Options?.Count ?? 0}");
                            var attributeOptions = new Dictionary<string, int?>();
                            foreach (var o in su.OptionSet.Options)
                            {
                                var label = o.Label?.UserLocalizedLabel?.Label ?? string.Empty;
                                attributeOptions[label] = o.Value.Value;
                            }
                            allOptionsData[su.LogicalName] = attributeOptions;
                        }
                    }
                    tracingService.Trace($"Finished processing all attributes. Total attributes with options: {allOptionsData.Count}");
                    optionsData = allOptionsData;
                }
                else
                {
                    tracingService.Trace($"Retrieving single attribute: {logicalName}");
                    var req = new RetrieveAttributeRequest
                    {
                        EntityLogicalName = entityLogicalName,
                        LogicalName = logicalName
                    };

                    tracingService.Trace($"Sending RetrieveAttributeRequest using values EntityLogicalName:{entityLogicalName}, LogicalName:{logicalName}");
                    var response = (RetrieveAttributeResponse)service.Execute(req);
                    tracingService.Trace($"RetrieveAttributeRequest call succeeded. AttributeType: {response.AttributeMetadata?.AttributeType}");

                    var singleAttributeOptions = new Dictionary<string, int?>();
                    switch (response.AttributeMetadata)
                    {
                        case BooleanAttributeMetadata b:
                            tracingService.Trace($"Processing Boolean attribute: {b.LogicalName}");
                            var trueLabel = b.OptionSet.TrueOption.Label?.UserLocalizedLabel?.Label ?? "True";
                            var falseLabel = b.OptionSet.FalseOption.Label?.UserLocalizedLabel?.Label ?? "False";
                            singleAttributeOptions[trueLabel] = b.OptionSet.TrueOption.Value;
                            singleAttributeOptions[falseLabel] = b.OptionSet.FalseOption.Value;
                            break;
                        case MultiSelectPicklistAttributeMetadata m:
                            tracingService.Trace($"Processing MultiSelectPicklist attribute: {m.LogicalName}, Options count: {m.OptionSet?.Options?.Count ?? 0}");
                            foreach (var o in m.OptionSet.Options)
                            {
                                var label = o.Label?.UserLocalizedLabel?.Label ?? string.Empty;
                                singleAttributeOptions[label] = o.Value.Value;
                            }
                            break;
                        case PicklistAttributeMetadata p:
                            tracingService.Trace($"Processing Picklist attribute: {p.LogicalName}, Options count: {p.OptionSet?.Options?.Count ?? 0}");
                            foreach (var o in p.OptionSet.Options)
                            {
                                var label = o.Label?.UserLocalizedLabel?.Label ?? string.Empty;
                                singleAttributeOptions[label] = o.Value.Value;
                            }
                            break;
                        case StateAttributeMetadata se:
                            tracingService.Trace($"Processing State attribute: {se.LogicalName}, Options count: {se.OptionSet?.Options?.Count ?? 0}");
                            foreach (var o in se.OptionSet.Options)
                            {
                                var label = o.Label?.UserLocalizedLabel?.Label ?? string.Empty;
                                singleAttributeOptions[label] = o.Value.Value;
                            }
                            break;
                        case StatusAttributeMetadata su:
                            tracingService.Trace($"Processing Status attribute: {su.LogicalName}, Options count: {su.OptionSet?.Options?.Count ?? 0}");
                            foreach (var o in su.OptionSet.Options)
                            {
                                var label = o.Label?.UserLocalizedLabel?.Label ?? string.Empty;
                                singleAttributeOptions[label] = o.Value.Value;
                            }
                            break;
                        default:
                            throw new InvalidPluginExecutionException($"The {logicalName} attribute doesn't have options.");
                    }
                    tracingService.Trace($"Finished processing single attribute. Total options collected: {singleAttributeOptions.Count}");
                    optionsData = singleAttributeOptions;
                }

                var jsonOptions = JsonSerializer.Serialize(optionsData);
                tracingService.Trace($"Serialized options to JSON. Length: {jsonOptions.Length}");
                context.OutputParameters["Options"] = jsonOptions;
                tracingService.Trace("RetrieveOptions plugin execution completed successfully.");
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException($"An error occurred in sample_RetrieveOptions: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                tracingService.Trace("sample_RetrieveOptions: {0}", ex.ToString());
                throw;
            }
        }
    }
}
