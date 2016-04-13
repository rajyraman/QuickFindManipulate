using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace RYR.QuickFindManipulate
{
    public class QuickFindPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            try
            {
                if (!context.InputParameters.Contains("Query")) return;

                var query = context.InputParameters["Query"] as QueryExpression;
                if (query == null)
                {
                    var fetchXml = context.InputParameters["Query"] as FetchExpression;

                    query = ((FetchXmlToQueryExpressionResponse)service.Execute(new FetchXmlToQueryExpressionRequest { FetchXml = fetchXml.Query })).Query;

                    context.InputParameters["Query"] = query;
                }

                var likeValue = string.Empty;
                var likeAttribute = string.Empty;
                var primaryKeyFilter = new FilterExpression(LogicalOperator.Or);
                var likeValueGuid = Guid.Empty;

                foreach (var filter in query.Criteria.Filters.ToList())
                {
                    if (filter.IsQuickFindFilter)
                    {
                        foreach (var condition in filter.Conditions.ToList())
                        {
                            if (condition.Operator == ConditionOperator.Like)
                            {
                                likeValue = condition.Values?[0].ToString();
                                likeAttribute = condition.AttributeName;
                                if (!string.IsNullOrEmpty(likeValue) && likeValue.EndsWith("%"))
                                {
                                    likeValue = likeValue.Substring(0, likeValue.Length - 1);
                                    var attributeDetail = RetrieveAttributeDetail(service, query.EntityName, condition.AttributeName);

                                    if (Guid.TryParse(likeValue, out likeValueGuid))
                                    {
                                        primaryKeyFilter.AddCondition(attributeDetail.Item2, ConditionOperator.Equal, likeValue);
                                        filter.Conditions.Remove(condition);
                                    }
                                    else
                                    {
                                        likeValue = string.Empty;
                                    }
                                }
                                else
                                {
                                    likeValue = string.Empty;
                                }
                            }
                        }
                        if (filter.Conditions.Count == 0)
                        {
                            query.Criteria.Filters.Remove(filter);
                        }
                    }
                }
                if (likeValue != string.Empty)
                {
                    primaryKeyFilter.AddCondition(RetrievePrimaryKeyAttributeName(service, query.EntityName), ConditionOperator.Equal, likeValue);
                    query.Criteria.AddFilter(primaryKeyFilter);
                }
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }

        private Tuple<bool, string> RetrieveAttributeDetail(IOrganizationService service, string entityName, string attributeName)
        {
            var attributeMetadata = (RetrieveAttributeResponse)service.Execute(
                new RetrieveAttributeRequest
                {
                    EntityLogicalName = entityName,
                    RetrieveAsIfPublished = true,
                    LogicalName = attributeName
                });
            if (!string.IsNullOrEmpty(attributeMetadata.AttributeMetadata.AttributeOf))
            {
                attributeMetadata = (RetrieveAttributeResponse)service.Execute(
                    new RetrieveAttributeRequest
                    {
                        EntityLogicalName = entityName,
                        RetrieveAsIfPublished = true,
                        LogicalName = attributeMetadata.AttributeMetadata.AttributeOf
                    });
            }
            var isValidForGuid = attributeMetadata.AttributeMetadata.AttributeType.Value == Microsoft.Xrm.Sdk.Metadata.AttributeTypeCode.Lookup ||
                attributeMetadata.AttributeMetadata.AttributeType.Value == Microsoft.Xrm.Sdk.Metadata.AttributeTypeCode.Customer ||
                attributeMetadata.AttributeMetadata.AttributeType.Value == Microsoft.Xrm.Sdk.Metadata.AttributeTypeCode.Uniqueidentifier;

            return new Tuple<bool, string>(isValidForGuid, attributeMetadata.AttributeMetadata.LogicalName);
        }

        private string RetrievePrimaryKeyAttributeName(IOrganizationService service, string entityName)
        {
            var entityMetadata = (RetrieveEntityResponse)service.Execute(
                new RetrieveEntityRequest {
                    LogicalName = entityName,
                    EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Entity });
            return entityMetadata.EntityMetadata.PrimaryIdAttribute;
        }
    }
}
