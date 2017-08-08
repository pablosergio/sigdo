using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Elfec.Sigdo
{
    public class ServerObjectModelUtils
    {
        private SPSite oSPSite;

        public ServerObjectModelUtils(SPSite site) {
            oSPSite = site;
        }

        public void CreateCustomContentTypes(string nameCustomType, string nameGroupContentType, List<SPField> siteColumns)
        {
            using (SPWeb oSPWeb = oSPSite.RootWeb)
            {

                // Get a reference to the Document or Item content type.
                SPContentType parentCType = oSPWeb.AvailableContentTypes[SPBuiltInContentTypeId.Document];

                // Create a Customer content type derived from the Item content type.
                SPContentType childCType = new SPContentType(parentCType, oSPWeb.ContentTypes, nameCustomType);

                childCType.Group = nameGroupContentType;

                // Add the new content type to the site collection.
                childCType = oSPWeb.ContentTypes.Add(childCType);

                foreach (SPField field in siteColumns)
                {
                    SPFieldLink fieldLink = new SPFieldLink(field);
                    childCType.FieldLinks.Add(fieldLink);

                }
                string[] fieldsToHide = new string[] { "Título" };
                foreach (string fieldDispName in fieldsToHide)
                {
                    SPField field = childCType.Fields[fieldDispName];
                    childCType.FieldLinks[field.Id].Hidden = true;
                }

                childCType.Update();

                oSPWeb.Update();
            }
        }

        // Delete site columns for group
        public void DeleteCustomSiteColumns(string groupColumn)
        {
            using (SPWeb oSPWeb = oSPSite.RootWeb)
            {

                List<SPField> fieldsInGroup = new List<SPField>();
                SPFieldCollection allFields = oSPWeb.Fields;
                for (int i = 0; i < allFields.Count; i++)
                {
                    if (allFields[i].Group.Equals(groupColumn))
                    {
                        allFields[i].Delete();
                    }
                }
        
                oSPWeb.Update();
            }
            
        }

        // Get field or sitesite columns for group
        public List<SPField> GetCustomSiteColumns(string groupSiteColumnName, string[] fields)
        {
            string groupColumn = groupSiteColumnName;
            List<SPField> siteColumns = new List<SPField>();
            using (SPWeb oSPWeb = oSPSite.RootWeb)
            {

                List<SPField> fieldsInGroup = new List<SPField>();
                SPFieldCollection allFields = oSPWeb.Fields;
                foreach (SPField field in allFields)
                {
                    if (field.Group.Equals(groupColumn))
                    {
                        foreach (string nameColumn in fields)
                        {
                            if (nameColumn == field.StaticName)
                            {
                                siteColumns.Add(field);
                            }
                        }
                    }
                }
            }
            
            return siteColumns;

        }

        // Delete ContentType
        public void DeleteContentType(string nameContentType)
        {

            using (SPWeb oSPWeb = oSPSite.RootWeb)
            {
                //Delete existing content type
                oSPWeb.AllowUnsafeUpdates = true;
                SPContentType contentType = oSPWeb.ContentTypes.Cast<SPContentType>().FirstOrDefault(c => c.Name == nameContentType);

                if (contentType != null)
                {
                    contentType.Delete();
                }

                oSPWeb.Update();
            }
      
        }


        // Set Content Types to Document Library
        public void SetContentTypesToDocumentLibrary(string nameDocumentLibrary, string nameGroupContentType)
        {

                using (SPWeb oSPWeb = oSPSite.RootWeb)
                {
                    List<SPContentType> resultList = new List<SPContentType>();
                    SPContentTypeCollection allWebContentTypes = oSPWeb.ContentTypes;
                    foreach (SPContentType contentType in allWebContentTypes)
                    {
                        if (contentType.Group == nameGroupContentType)
                            resultList.Add(contentType);
                    }
                    SPDocumentLibrary library = oSPWeb.Lists[nameDocumentLibrary] as SPDocumentLibrary;
                    library.ContentTypesEnabled = true;
                    TryAddEachContentTypeToLibrary(library, resultList);
                    library.Update();

                }
            
        }

        // Set Content Types to Document Library
        public void SetContentTypesToDocumentLibrary(string url, string nameDocumentLibrary, string nameGroupContentType)
        {
            using (SPSite site = new SPSite(url))
            {
                using (SPWeb oSPWeb = site.RootWeb)
                {
                    List<SPContentType> resultList = new List<SPContentType>();
                    SPContentTypeCollection allWebContentTypes = oSPWeb.ContentTypes;
                    foreach (SPContentType contentType in allWebContentTypes)
                    {
                        if (contentType.Group == nameGroupContentType)
                            resultList.Add(contentType);
                    }
                    SPDocumentLibrary library = oSPWeb.Lists[nameDocumentLibrary] as SPDocumentLibrary;
                    library.ContentTypesEnabled = true;
                    TryAddEachContentTypeToLibrary(library, resultList);
                    library.Update();

                }    
            }
            

        }

        private void TryAddEachContentTypeToLibrary(SPDocumentLibrary docLibrary, List<SPContentType> ContentTypesOnThisWeb)
        {
            SPContentTypeCollection libraryContentTypes = docLibrary.ContentTypes;
            foreach (SPContentType contentType in ContentTypesOnThisWeb)
            {
                if (DocLibContainsContentType(libraryContentTypes, contentType))
                    continue;
                libraryContentTypes.Add(contentType);
            }
        }

        private bool DocLibContainsContentType(SPContentTypeCollection libraryContentTypes, SPContentType contentType)
        {
            foreach (SPContentType libraryContentType in libraryContentTypes)
            {
                if (libraryContentType.Name == contentType.Name)
                    return true;
            }
            return false;
        }

        public void CreateDocumentLibrary(string nameDocumentLibrary, string title, string description, bool quickLaunch, bool enableVersioning = false)
        {
            using (SPWeb web  = oSPSite.RootWeb) // will get the current web
            {
                //Check to see if list already exists
                try
                {
                    SPDocumentLibrary targetList = oSPSite.RootWeb.Lists[nameDocumentLibrary] as SPDocumentLibrary;

                }
                catch (Exception)
                {
                    Guid newListId = web.Lists.Add(
                        title, // List Title
                        description, // List Description
                        SPListTemplateType.DocumentLibrary // List Template
                        //docTemplate // Document Template (i.e. Excel)
                        );

                    SPDocumentLibrary newLibrary = web.Lists[newListId] as SPDocumentLibrary;
                    newLibrary.OnQuickLaunch = quickLaunch;
                    newLibrary.EnableVersioning = enableVersioning;

                    newLibrary.Update();
                }
                 
            
            }
           
        }

        public void CreateList(string listName, string title, string description, SPListTemplateType template, bool quickLaunch, int readSecurity, int writeSecurity, bool enableVersioning = false)
        {
            using (SPWeb oSPWeb = oSPSite.RootWeb)
            {
                //Check to see if list already exists
                try
                {
                    SPList targetList = oSPSite.RootWeb.Lists[listName];
                }
                catch (ArgumentException)
                {
                    //The list does not exist, thus you can create it    
                    //The list does not exist, thus you can create it
                    Guid listId = oSPWeb.Lists.Add(title,
                        description,
                        template
                   );

                    SPList newList = oSPWeb.Lists[listId];
                    newList.OnQuickLaunch = quickLaunch;
                    newList.EnableVersioning = enableVersioning;
                    newList.ReadSecurity = readSecurity; // All users have Read access to all items
                    newList.WriteSecurity = writeSecurity; // Users can modify only items that they created


                    newList.Update();

                }
            
            }
        }

        public void UpdateList(string listName, string fieldName, string description, SPFieldType fieldType, bool required)
        {
            using (SPWeb oSPWeb = oSPSite.RootWeb)
            {
                //Check to see if list already exists
                try
                {
                    SPList targetList = oSPSite.RootWeb.Lists[listName];
                    string fieldCode = targetList.Fields.Add(fieldName, fieldType, required);
                    targetList.Fields[fieldCode].Description = description;
                    targetList.Fields[fieldCode].Update();
                    targetList.Update();
                }
                catch (ArgumentException)
                {
                    
                }

            }
        }

       
        public void DeleteList(string listName) 
        {
            //Check to see if list already exists
            try
            {
                SPList targetList = oSPSite.RootWeb.Lists[listName];
                targetList.Delete();
            }
            catch (ArgumentException)
            {
                //The list does not exist, thus you can create it                 
            }
        }

        public void CreateSiteColumns()
        {
            string groupColumn = "Sistema Correspondencia";

            using (SPWeb oSPWeb = oSPSite.RootWeb)
            {
                //if (!oSPWeb.Fields.ContainsField("Codigo"))
                //{
                string codigoField = oSPWeb.Fields.Add("Codigo", SPFieldType.Text, false);
                SPFieldText Codigo = (SPFieldText)oSPWeb.Fields.GetFieldByInternalName(codigoField);
                Codigo.Group = groupColumn;
                Codigo.Update();
                //}
                
                //if (!oSPWeb.Fields.ContainsField("Referencia"))
                //{
                    string referenciaField = oSPWeb.Fields.Add("Referencia", SPFieldType.Text, false);
                    SPFieldText Referencia = (SPFieldText)oSPWeb.Fields.GetFieldByInternalName(referenciaField);
                    Referencia.Group = groupColumn;
                    Referencia.Update();
                //}

                //if (!oSPWeb.Fields.ContainsField("Comentario")) {
                    string comentarioField = oSPWeb.Fields.Add("Comentario", SPFieldType.Text, false);
                    SPFieldText Comentario = (SPFieldText)oSPWeb.Fields.GetFieldByInternalName(comentarioField);
                    Comentario.Group = groupColumn;
                    Comentario.Update();
                //}
                
                
                //if (!oSPWeb.Fields.ContainsField("Remitente"))
                //{
                    string remitenteField = oSPWeb.Fields.Add("Remitente", SPFieldType.Text, false);
                    SPFieldText Remitente = (SPFieldText)oSPWeb.Fields.GetFieldByInternalName(remitenteField);
                    Remitente.Group = groupColumn;
                    Remitente.Update();

                //}
                
                //if (!oSPWeb.Fields.ContainsField("Factura")) {
                    string facturaField = oSPWeb.Fields.Add("Factura", SPFieldType.Number, false);
                    SPFieldNumber Factura = (SPFieldNumber)oSPWeb.Fields.GetFieldByInternalName(facturaField);
                    Factura.Group = groupColumn;
                    Factura.Update();
                //}
                
                
                //if (!oSPWeb.Fields.ContainsField("Fecha Receopcion")) {
                    string fechaRecepcionField = oSPWeb.Fields.Add("Fecha Recepcion", SPFieldType.DateTime, true);
                    SPFieldDateTime FechaRecepcion = (SPFieldDateTime)oSPWeb.Fields.GetFieldByInternalName(fechaRecepcionField);
                    FechaRecepcion.Group = groupColumn;
                    FechaRecepcion.Update();

                //}
                
                //Method 3 using Field schema
                //if (!oSPWeb.Fields.ContainsField("Nota"))
                //{
                    string notaField = oSPWeb.Fields.Add("Nota", SPFieldType.Note, false);
                    SPFieldMultiLineText Nota = (SPFieldMultiLineText)oSPWeb.Fields.GetFieldByInternalName(notaField);
                    Nota.Group = groupColumn;
                    Nota.Update();
                //}

                //if (!oSPWeb.Fields.ContainsField("Gerencias"))
                //{

                    // Lookup Column
                    SPList gerenciasList = oSPWeb.Lists.TryGetList("Gerencias");

                    string GerenciaField = oSPWeb.Fields.AddLookup("Gerencia", gerenciasList.ID, true);
                    SPFieldLookup gerenciaLookup = (SPFieldLookup)oSPWeb.Fields.GetFieldByInternalName(GerenciaField);

                    gerenciaLookup.Group = groupColumn;
                    gerenciaLookup.LookupField = "Title";
                    gerenciaLookup.Update();
                //}

                SPList listas = oSPWeb.Lists.TryGetList("Listas");
                //if (!oSPWeb.Fields.ContainsField("Nivel Prioridad"))
                //{

                    string nivelPrioridadField = oSPWeb.Fields.AddLookup("Nivel Prioridad", listas.ID, true);
                    SPFieldLookup nivelPrioridadLookup = (SPFieldLookup)oSPWeb.Fields.GetFieldByInternalName(nivelPrioridadField);
                    nivelPrioridadLookup.Group = groupColumn;
                    nivelPrioridadLookup.LookupField = "Title";
                    nivelPrioridadLookup.Update();
                //}

                //if (!oSPWeb.Fields.ContainsField("Clasificacion"))
                //{

                    string clasificacionField = oSPWeb.Fields.AddLookup("Clasificacion", listas.ID, true);
                    SPFieldLookup clasificacionLookup = (SPFieldLookup)oSPWeb.Fields.GetFieldByInternalName(clasificacionField);
                    clasificacionLookup.Group = groupColumn;
                    clasificacionLookup.LookupField = "Title";
                    clasificacionLookup.Update();
                //}

                //if (!oSPWeb.Fields.ContainsField("Estado"))
                //{

                    string estadoField = oSPWeb.Fields.AddLookup("Estado", listas.ID, false);
                    SPFieldLookup estadoLookup = (SPFieldLookup)oSPWeb.Fields.GetFieldByInternalName(estadoField);
                    estadoLookup.Group = groupColumn;
                    estadoLookup.LookupField = "Title";
                    estadoLookup.Update();
                //}
            }
        }

        public void CreateWebSite(string url, string title, string description, string siteTemplate)
        {

            using (SPWeb web = oSPSite.OpenWeb(oSPSite.RootWeb.ID))
            {
                try
                {
                    web.AllowUnsafeUpdates = true;

                  
                    // Site creation with unique permissions
                    SPWebCollection webs = web.Webs;
                    SPWeb newWeb = webs.Add(url, title, description, 1033, siteTemplate, true, false);

                    // Owners, members and visitors groups creation
                    SPGroup owners = SPGroupsHelper.AddGroup(newWeb, SPGroupsHelper.AssociatedGroupTypeEnum.Owners);
                    SPGroup members = SPGroupsHelper.AddGroup(newWeb, SPGroupsHelper.AssociatedGroupTypeEnum.Members);
                    SPGroup visitors = SPGroupsHelper.AddGroup(newWeb, SPGroupsHelper.AssociatedGroupTypeEnum.Visitors);

                    // Changing the request access email to current user
                    newWeb.RequestAccessEmail = newWeb.CurrentUser.Email;

                    // Save changes
                    newWeb.Update();

                    // Disposing new web object
                    newWeb.Dispose();
       
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    web.AllowUnsafeUpdates = false;
                }
            }
           
        }

        

    }

    public static class SPGroupsHelper
    {
 
        public enum AssociatedGroupTypeEnum
        {
            Owners,
            Members,
            Visitors
        };
 
        public static SPGroup AddGroup(SPWeb web, AssociatedGroupTypeEnum associateGroupType)
        {
            switch (associateGroupType)
            {
                case AssociatedGroupTypeEnum.Owners:
                    return AddGroup(web, "{0} Owners", "Use this group to give people full control permissions to the SharePoint site: {0}", SPRoleType.Administrator, "{0} Owners");
                case AssociatedGroupTypeEnum.Members:
                    return AddGroup(web, "{0} Members", "Use this group to give people contribute permissions to the SharePoint site: {0}", SPRoleType.Contributor,"{0} Owners");
                case AssociatedGroupTypeEnum.Visitors:
                    return AddGroup(web, "{0} Vistors", "Use this group to give people read permissions to the SharePoint site: {0}", SPRoleType.Reader,"{0} Owners");
                default:
                    return null;
            }
        }
 
        public static SPGroup AddGroup(SPWeb web, string groupNameFormatString, string descriptionFormatString, SPRoleType roleType, string ownerNameFormatString)
        {
            web.SiteGroups.Add(string.Format(groupNameFormatString, web.Title), web.CurrentUser, web.CurrentUser, string.Format(descriptionFormatString, web.Name));
 
            SPGroup group = web.SiteGroups[string.Format(groupNameFormatString,web.Title)];
            try
            {
                SPGroup owner = web.SiteGroups[string.Format(ownerNameFormatString, web.Title)];
                group.Owner = owner;
            }
            catch { }
 
            if (descriptionFormatString.IndexOf("{0}") != -1)
            {
                SPListItem item = web.SiteUserInfoList.GetItemById(group.ID);
                item["Notes"] = string.Format(descriptionFormatString, string.Format("<a href=\"{0}\">{1}</a>", web.Url, web.Name));
                item.Update();
            }
 
            SPRoleAssignment roleAssignment = new SPRoleAssignment(group);
            roleAssignment.RoleDefinitionBindings.Add(web.RoleDefinitions.GetByType(roleType));
            web.RoleAssignments.Add(roleAssignment);
            switch(roleType)
            {
                case SPRoleType.Administrator:
                    group.AllowMembersEditMembership = false;
                    group.OnlyAllowMembersViewMembership = true;
                    group.AllowRequestToJoinLeave = false;
                    group.AutoAcceptRequestToJoinLeave = false;
                    web.AssociatedOwnerGroup = group;
 
                    break;
                case SPRoleType.Contributor:
                    group.AllowMembersEditMembership = false;
                    group.AllowRequestToJoinLeave = false;
                    group.AutoAcceptRequestToJoinLeave = false;
                    group.OnlyAllowMembersViewMembership = false;
                    web.AssociatedMemberGroup = group;
                    break;
                case SPRoleType.Reader:
                    group.AllowMembersEditMembership = false;
                    group.OnlyAllowMembersViewMembership = true;
                    group.AllowRequestToJoinLeave = false;
                    group.AutoAcceptRequestToJoinLeave = false;
                    web.AssociatedVisitorGroup = group;
                    break;
            }
            group.Update();
            web.Update();
            return group;
        }
    }
}
