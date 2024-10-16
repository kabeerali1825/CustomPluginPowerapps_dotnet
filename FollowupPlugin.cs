using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Custom_Plugin_PM
{
    public class PM_Plugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.  
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.  
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.  
                Entity entity = (Entity)context.InputParameters["Target"];

                // Obtain the IOrganizationService instance which you will need for  
                // web service calls.  
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                // Check if the "projectmanager" field is being updated
                if (entity.Contains("proj_projectmanager"))
                {
                    tracingService.Trace("I am in the If Block");
                    // Retrieve the Pre-Image (contains the old value)
                    Entity preImage = (Entity)context.PreEntityImages["PreImage"];
                    tracingService.Trace("Pre Image"+preImage);
                    Guid oldProjectManagerId = preImage.GetAttributeValue<EntityReference>("proj_projectmanager")?.Id ?? Guid.Empty;
                    tracingService.Trace("After Creating pre image ");
                    // Query the resource entity to get the associated user ID
                    QueryExpression query = new QueryExpression("proj_resources") // Assuming 'resource' is the logical name of your resource entity
                    {
                        ColumnSet = new ColumnSet("proj_associateduser") // Specify the fields you need
                    };
                    query.Criteria.AddCondition("proj_resourcesid", ConditionOperator.Equal, oldProjectManagerId); // Adjust if your primary key is different
                    tracingService.Trace("After getting Old PM "+ oldProjectManagerId);
                  
                    EntityCollection results = service.RetrieveMultiple(query);
                    tracingService.Trace("After Retreving from query");
                    Guid userId;
                    Guid newPMUserID;
                    if (results.Entities.Count > 0)
                    {
                        tracingService.Trace("I am in the results if block");
                        Entity resource = results.Entities[0];
                        tracingService.Trace("after entity resource");
                        //userId = resource.GetAttributeValue<Guid>("proj_associateduser");
                        EntityReference associatedUserRef = resource.GetAttributeValue<EntityReference>("proj_associateduser");
                        userId = associatedUserRef?.Id ?? Guid.Empty;  // Use the Id from the EntityReference or set it to Guid.Empty if null

                        tracingService.Trace("This is user ID OF OLD PM" + userId);
                        // Retrieve the Post-Image (contains the new value)
                        Entity postImage = (Entity)context.PostEntityImages["PostImage"];
                        Guid newProjectManagerId = postImage.GetAttributeValue<EntityReference>("proj_projectmanager")?.Id ?? Guid.Empty;
                        QueryExpression queryfornewPM = new QueryExpression("proj_resources") // Assuming 'resource' is the logical name of your resource entity
                        {
                            ColumnSet = new ColumnSet("proj_associateduser") // Specify the fields you need
                        };
                        queryfornewPM.Criteria.AddCondition("proj_resourcesid", ConditionOperator.Equal, newProjectManagerId); // Adjust if your primary key is different

                        EntityCollection resultsforNewPM = service.RetrieveMultiple(queryfornewPM);

                        if (resultsforNewPM.Entities.Count > 0)
                        {
                            Entity resourceForNewPM = resultsforNewPM.Entities[0];
                            //newPMUserID = resourceForNewPM.GetAttributeValue<Guid>("proj_associateduser");
                            EntityReference associatedUserRefNewPM = resourceForNewPM.GetAttributeValue<EntityReference>("proj_associateduser");
                            newPMUserID= associatedUserRefNewPM?.Id ?? Guid.Empty;


                            // Unshare the record with the old Project Manager
                            if (userId != Guid.Empty)
                            {
                                try
                                {
                                    // Unshare the record
                                    var revokeRequest = new RevokeAccessRequest
                                    {
                                        Target = new EntityReference(entity.LogicalName, entity.Id),
                                        Revokee = new EntityReference("systemuser", userId)
                                    };
                                    service.Execute(revokeRequest);
                                    tracingService.Trace($"Record unshared with old Project Manager: {userId}");
                                }
                                catch (Exception ex)
                                {
                                    tracingService.Trace($"Failed to unshare the record: {ex.Message}");
                                }
                            }

                            // Share the record with the new Project Manager
                            if (newPMUserID != Guid.Empty)
                            {
                                try
                                {
                                    var shareRequest = new GrantAccessRequest
                                    {
                                        Target = new EntityReference(entity.LogicalName, entity.Id),
                                        PrincipalAccess = new PrincipalAccess
                                        {
                                            Principal = new EntityReference("systemuser", newPMUserID),
                                            AccessMask = AccessRights.ReadAccess | AccessRights.WriteAccess // Customize access as needed
                                        }
                                    };
                                    service.Execute(shareRequest);
                                    tracingService.Trace($"Record shared with new Project Manager: {newPMUserID}");
                                }
                                catch (Exception ex)
                                {
                                    tracingService.Trace($"Failed to share the record: {ex.Message}");
                                }
                            }
                        }
                    }

                    
                    
                }

            }
        }
    }
}
