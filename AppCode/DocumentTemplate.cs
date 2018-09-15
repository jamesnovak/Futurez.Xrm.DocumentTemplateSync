﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using Futurez.Xrm.DocumentTemplateSync;
using Futurez.Xrm.Tools;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Futurez.Entities
{
    [DefaultProperty("Name")]
    public partial class DocumentTemplateEdit
    {
        public DocumentTemplateEdit(Entity template)
        {
            Id = template.Id;
            Name = template.GetAttribValue<string>("name");
            Description = template.GetAttribValue<string>("description");
            Type = template.GetFormattedAttribValue("documenttype");
            TypeValue = template.GetAttribValue<OptionSetValue>("documenttype").Value;
            AssociatedEntity = template.GetFormattedAttribValue("associatedentitytypecode");
            AssociatedEntityLogicalName = template.GetAttribValue<string>("associatedentitytypecode");

            Status = template.GetFormattedAttribValue("status");

            var entityRef = template.GetAttribValue<EntityReference>("createdby");
            CreatedBy = (entityRef != null) ? entityRef.Name : null;

            var dt = template.GetAttribValue<DateTime?>("createdon");
            if (dt.HasValue) {
                CreatedOn = dt.Value.ToLocalTime();
            }

            entityRef = template.GetAttribValue<EntityReference>("modifiedby");
            ModifiedBy = (entityRef != null) ? entityRef.Name : null;

            dt = template.GetAttribValue<DateTime?>("modifiedon");
            if (dt.HasValue) {
                ModifiedOn = dt.Value.ToLocalTime();
            }

            EntityLogicalName = template.LogicalName;
            TemplateScope = (template.LogicalName == "documenttemplate") ? "System" : "Personal";

            var content = template.GetAttribValue<string>("content");
            Content = (content == null) ? null : Convert.FromBase64String(content);

            var ext = TypeValue == 1 ? Constants.FILE_EXT_EXCEL : Constants.FILE_EXT_WORD;

            // make sure we have unique names
            FileName = Name + ext;
        }
        #region Attributes
        [Description("Entity Logical Name")]
        [Category("General")]
        [Browsable(false)]
        [DisplayName("Entity Logical Name")]
        public string EntityLogicalName { get; private set; }

        [Description("Document Template Id")]
        [Category("General")]
        [Browsable(false)]
        public Guid Id { get; private set; }

        [Description("Document Template Name")]
        [Category("General")]
        public string Name { get; private set; }

        [Description("Document Template Filename")]
        [Category("General")]
        public string FileName { get; private set; }

        [Description("Description for this Document Template")]
        [Category("General")]
        [EditorAttribute("System.ComponentModel.Design.MultilineStringEditor, System.Design", "System.Drawing.Design.UITypeEditor")]
        public string Description { get; private set; }

        [Description("Current Document Template Status")]
        [Category("Locked")]
        public string Status { get; private set; }

        [DisplayName("Associated Entity SchemaName")]
        [Description("Entity SchemaName to which this Document Template is assocaited")]
        [Category("Locked")]
        [Browsable(false)]
        public string AssociatedEntityLogicalName { get; private set; }

        [DisplayName("Associated Entity Name")]
        [Description("Name of the Entity to which this Document Template is assocaited")]
        [Category("Locked")]
        public string AssociatedEntity { get; private set; }

        [DisplayName("Created On")]
        [Description("Date/Time on which this Document Template was created")]
        [Category("Locked")]
        public DateTime? CreatedOn { get; private set; }

        [DisplayName("Created By")]
        [Description("System user who created this Document Template")]
        [Category("Locked")]
        public string CreatedBy { get; private set; }

        [DisplayName("Modified On")]
        [Description("Date/Time on which this Document Template was last modified")]
        [Category("Locked")]
        public DateTime? ModifiedOn { get; private set; }

        [DisplayName("Modified By")]
        [Description("System user who last modified this Document Template")]
        [Category("Locked")]
        public string ModifiedBy { get; private set; }

        [Description("Document Template Type")]
        [DisplayName("Content Type")]
        [Category("Locked")]
        public string Type { get; private set; }

        [Description("Document Template Type")]
        [DisplayName("Content Type")]
        [Category("Locked")]
        [Browsable(false)]
        public int TypeValue { get; private set; }

        [Description("Language setting for this Document Template")]
        [Category("Locked")]
        [Browsable(false)]
        public string Language { get; private set; }

        [Description("System or Personal document template")]
        [Category("General")]
        [DisplayName("Template Scope")]
        public string TemplateScope { get; private set; }

        [Description("Document Content")]
        [Category("General")]
        [Browsable(false)]
        [DisplayName("Document Content")]
        public byte[] Content { get; private set; }

        #endregion

        #region Helper Methods 
        /// <summary>
        /// Helper method to return document templates using several filter options 
        /// </summary>
        /// <param name="service"></param>
        /// <param name="syncSettings"></param>
        /// <param name="includeContent"></param>
        /// <param name="templateIds"></param>
        /// <returns></returns>
        public static List<DocumentTemplateEdit> GetDocumentTemplates(IOrganizationService service, SyncSettings syncSettings = null, bool includeContent = false, List<Guid> templateIds = null)
        {
            var templates = new List<DocumentTemplateEdit>();

            var query = new QueryExpression() {
                EntityName = "documenttemplate",
                ColumnSet = new ColumnSet("name", "documenttype", "documenttemplateid", "associatedentitytypecode", "status"),
                Criteria = new FilterExpression(LogicalOperator.And)
            };

            if (includeContent) {
                query.ColumnSet.AddColumn("content");
            }

            if (templateIds != null) {
                query.Criteria.AddCondition(new ConditionExpression("documenttemplateid", ConditionOperator.In, templateIds));
            }

            // apply filters from the sync settings file 
            if (syncSettings != null) 
                {
                if (syncSettings.ActiveOnly) {
                    query.Criteria.AddCondition(new ConditionExpression("status", ConditionOperator.Equal, false));
                }

                if (syncSettings.DocumentTypes == SelectedDocumentTypes.Excel) {
                    query.Criteria.AddCondition(new ConditionExpression("documenttype", ConditionOperator.Equal, 1));
                }
                else if (syncSettings.DocumentTypes == SelectedDocumentTypes.Word) {
                    query.Criteria.AddCondition(new ConditionExpression("documenttype", ConditionOperator.Equal, 2));
                }

                if (syncSettings.EntityTypeFilters != null && syncSettings.EntityTypeFilters.Count > 0) {
                    query.Criteria.AddCondition(new ConditionExpression("associatedentitytypecode", ConditionOperator.In, syncSettings.EntityTypeFilters));
                }

                if (syncSettings.FileNameFilters != null && syncSettings.FileNameFilters.Count > 0) {
                    foreach (var filter in syncSettings.FileNameFilters) {
                        query.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.NotLike, $"%{filter}%"));
                    }
                }
            }

            // grab the entities and add them to the doc template collecion 
            var results = service.RetrieveMultiple(query);

            foreach (var entity in results.Entities) {
                templates.Add(new DocumentTemplateEdit(entity));
            }

            return templates;
        }

        #endregion
    }
}
