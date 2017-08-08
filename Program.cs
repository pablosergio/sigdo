using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Elfec.Sigdo.Install
{
    class Program
    {
        static void Main(string[] args)
        {
            var solutionFileName = @"SolutionFiles\Elfec.Sigdo.wsp";
            var solutionId = "b35d3579-1e9d-400d-974b-a52503076369";
            string[] webApplicationNames = new string[] { "hostdns" };

            var solutionCommand = new SolutionCommand(solutionFileName, solutionId);
            solutionCommand.Execute();
            solutionCommand.Deploy(webApplicationNames);

            var siteURL = @"http://hostdns/";
            //var siteColumnFeatureId = "71b2c02a-b3d0-4a85-8b08-5af4b20f23dd";
            //var featureCommand = new FeatureCommnad(siteURL, siteColumnFeatureId);
            //featureCommand.Execute();
            //featureWebCorrespondencia.Rollback();
            //CreateNewWebSite();
            //CreateDocumentLibrary("Correspondencia", "Correspondencia", "Correspondencia Recibida");
            //DeleteCustomSiteColumns("Sistema Correspondencia");
        }
        // Delete site columns for group
        private static void DeleteCustomSiteColumns(string groupSiteColumnName)
        {
            string groupColumn = groupSiteColumnName;
            using (SPSite oSPSite = new SPSite("http://hostdns/"))
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
                    /*foreach (SPField field in allFields)
                    {
                        if (field.Group.Equals(groupColumn))
                        {
                            field.Delete();
                        }
                    }*/

                    oSPWeb.Update();
                }
            }

        }

    }
}
