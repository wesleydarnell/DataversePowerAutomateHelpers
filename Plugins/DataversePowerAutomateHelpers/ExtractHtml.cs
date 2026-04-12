using Microsoft.Xrm.Sdk;
using System;
using System.Linq;
using System.ServiceModel;
using System.Text;
using HtmlAgilityPack;

namespace DataversePowerAutomateHelpers
{
    /// <summary>
    /// Provides logic implementation for the HtmlExtractor Custom API
    /// </summary>
    public class ExtractHtml : IPlugin
    {
        /*
        Input Parameters:
        | String  | HtmlString                         | Required | Input stringified Html
        | String  | XPathSelector                      | Required | The XPath selector for the intended element (e.g. //div[@class='ck-content']). If the selector is not found, the original Html string will be returned.
        | Boolean | ReturnSelectorElementAndChildren   | Required | Default false. If true, the element identified by the selector and its children will be returned. If false, only the inner content of the selector element is returned, without the parent itself.

        Output Parameters:
        | Boolean | Success    | Indicates whether the operation was successful
        | String  | HtmlString | The resulting Html string after processing with the provided selector. If the selector is not found, the original Html string will be returned.
        */

        public void Execute(IServiceProvider serviceProvider)
        {
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            tracingService.Trace("ExtractHtml plugin execution started.");

            try
            {
                //Read input parameters
                tracingService.Trace("Reading input parameters...");

                if (!context.InputParameters.Contains("HtmlString") || string.IsNullOrEmpty((string)context.InputParameters["HtmlString"]))
                {
                    throw new InvalidPluginExecutionException("HtmlString input parameter is required.");
                }

                if (!context.InputParameters.Contains("XPathSelector") || string.IsNullOrEmpty((string)context.InputParameters["XPathSelector"]))
                {
                    throw new InvalidPluginExecutionException("XPathSelector input parameter is required.");
                }

                string htmlString = (string)context.InputParameters["HtmlString"];
                string xpathSelector = (string)context.InputParameters["XPathSelector"];
                bool returnSelectorElement = context.InputParameters.Contains("ReturnSelectorElementAndChildren")
                    && (bool)context.InputParameters["ReturnSelectorElementAndChildren"];

                tracingService.Trace($"Input - XPathSelector: '{xpathSelector}', ReturnSelectorElementAndChildren: {returnSelectorElement}");

                //Load the HTML document
                var htmlDoc = new HtmlDocument();
                htmlDoc.LoadHtml(htmlString);
                tracingService.Trace("HTML document loaded successfully.");

                //Select the first matching node using the XPath selector
                var selectedNode = htmlDoc.DocumentNode.SelectSingleNode(xpathSelector);

                if (selectedNode == null)
                {
                    //Selector not found, return original HTML
                    tracingService.Trace($"No element found matching selector '{xpathSelector}'. Returning original HTML.");
                    context.OutputParameters["Success"] = false;
                    context.OutputParameters["HtmlString"] = htmlString;
                }
                else
                {
                    tracingService.Trace($"Element found matching selector '{xpathSelector}'.");

                    string result;

                    if (returnSelectorElement)
                    {
                        //Return the matched element and all its children
                        result = selectedNode.OuterHtml;
                        tracingService.Trace("Returning selector element and children (OuterHtml).");
                    }
                    else
                    {
                        //Return only the inner content of the matched element
                        result = selectedNode.InnerHtml;
                        tracingService.Trace("Returning children only (InnerHtml).");
                    }

                    context.OutputParameters["Success"] = true;
                    context.OutputParameters["HtmlString"] = result;
                    tracingService.Trace($"Result HTML length: {result.Length}");
                }

                tracingService.Trace("ExtractHtml plugin execution completed successfully.");
            }

            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException($"An error occurred in sample_ExtractHtml: {ex.Message}", ex);
            }

            catch (Exception ex)
            {
                tracingService.Trace("sample_ExtractHtml: {0}", ex.ToString());
                throw;
            }
        }
    }
}
