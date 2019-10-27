using System;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using System.Text;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;
using System.Linq;
using System.Text.RegularExpressions;
using System.Runtime.Serialization.Json;
using System.Runtime.Serialization;
using System.IO;

namespace ItAintBoring.InspectionDemo
{

    [DataContract]
    internal class InspectionResults
    {
        [DataMember]
        public string ProductLength = "N/A";

        [DataMember]
        public string ProductWidth = "N/A";

        [DataMember]
        public string ProductShape = "N/A";

    }


    public class InspectionPlugin: IPlugin
    {

        private OptionSetValue passOption = new OptionSetValue(556780000);
        private OptionSetValue failOption = new OptionSetValue(556780001);
        private OptionSetValue naOption = new OptionSetValue(556780002);


        public void CreateInspectionItem(IOrganizationService service, string name, string result, Entity inspection)
        {
            Entity item = new Entity("ita_inspectionitem");
            item["ita_name"] = name;
            item["ita_result"] = (result == "Fail" ? failOption : (result == "Pass" ? passOption : naOption));
            item["ita_inspection"] = inspection.ToEntityReference();
            service.Create(item);
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService =
                (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.Depth > 1) return;
            if(context.MessageName == "Update" || context.MessageName == "Create")
            {
                var target = (Entity)context.InputParameters["Target"];
                if(target.Contains("ita_json"))
                {
                    string json = (string)target["ita_json"];
                    try
                    {
                        var ser = new DataContractJsonSerializer(typeof(InspectionResults));
                        var stream1 = new MemoryStream();
                        StreamWriter sw = new StreamWriter(stream1);
                        sw.Write(json);
                        sw.Flush();
                        stream1.Position = 0;
                        InspectionResults ir = (InspectionResults)ser.ReadObject(stream1);

                        QueryExpression qe = new QueryExpression("ita_inspectionitem");
                        qe.Criteria.AddCondition(new ConditionExpression("ita_inspection", ConditionOperator.Equal, target.Id));
                        qe.ColumnSet = new ColumnSet("ita_inspectionitemid");
                        var results = service.RetrieveMultiple(qe).Entities;
                        foreach (var e in results)
                        {
                            service.Delete(e.LogicalName, e.Id);
                        }
                        CreateInspectionItem(service, "Length", ir.ProductLength, target);
                        CreateInspectionItem(service, "Width", ir.ProductWidth, target);
                        CreateInspectionItem(service, "Shape", ir.ProductShape, target);
                    }
                    catch(Exception ex)
                    {
                        target["ita_json"] = ex.Message;
                        service.Update(target);
                    }
                }
            }
            else if (context.MessageName == "RetrieveMultiple")
            {
                if (context.Stage == 20)
                {
                    if (context.InputParameters.Contains("Query"))
                    {

                        FetchExpression query = null;
                        if (context.InputParameters["Query"] is QueryExpression)
                        {
                            var qe = (QueryExpression)context.InputParameters["Query"];
                            var conversionRequest = new QueryExpressionToFetchXmlRequest
                            {
                                Query = qe
                            };

                            var conversionResponse =
                                (QueryExpressionToFetchXmlResponse)service.Execute(conversionRequest);
                            query = new FetchExpression(conversionResponse.FetchXml);
                        }

                        if (context.InputParameters["Query"] is FetchExpression)
                        {
                            query = (FetchExpression)context.InputParameters["Query"];
                        }
                        string pattern = @"condition attribute=""ita_inspectionforproduct"" operator=""eq"" value=""(.*?)""";
                        var matches = Regex.Matches(query.Query, pattern, RegexOptions.Multiline);
                        if (matches.Count > 0 && matches[0].Groups.Count > 0)
                        {
                            context.SharedVariables["productnumber"] = matches[0].Groups[1].Value;
                        }
                    }
                }
                else
                {
                    EntityCollection col = (EntityCollection)context.OutputParameters["BusinessEntityCollection"];
                    if (context.SharedVariables.Contains("productnumber"))
                    {
                        string productNumber = (string)context.SharedVariables["productnumber"];
                        Entity entity = null;
                        if (productNumber == "333")
                        {
                            //Create a fake inspection record to pass the error back
                            entity = new Entity("ita_inspection");
                            entity.Id = Guid.NewGuid();
                            entity["ita_inspectionid"] = entity.Id;
                            entity["ita_inspectionforproduct"] = productNumber;
                            entity["ita_startinspectionerror"] = "Cannot inspect this product!";
                        }
                        else
                        {
                            //Get inspection record
                            //Will create here, but could also load an existing record
                            entity = new Entity("ita_inspection");
                            entity["ita_inspectionforproduct"] = productNumber;
                            entity["ita_name"] = productNumber;
                            entity.Id = service.Create(entity);
                            entity["ita_inspectionid"] = entity.Id;
                            
                            //Prepare JSON
                            InspectionResults ir = new InspectionResults();
                            if (productNumber == "111")
                            {
                                ir.ProductLength = "Fail";
                                ir.ProductWidth = "Pass";
                                ir.ProductShape = "Fail";
                            }
                            else
                            {
                                ir.ProductLength = "N/A";
                                ir.ProductWidth = "N/A";
                                ir.ProductShape = "N/A";
                            }
                            var stream1 = new MemoryStream();
                            var ser = new DataContractJsonSerializer(typeof(InspectionResults));
                            ser.WriteObject(stream1, ir);
                            stream1.Position = 0;
                            var sr = new StreamReader(stream1);
                            string json = sr.ReadToEnd();
                            entity["ita_json"] = json;

                        }
                        col.Entities.Add(entity);
                    }
                }
            }
        }
    }
}