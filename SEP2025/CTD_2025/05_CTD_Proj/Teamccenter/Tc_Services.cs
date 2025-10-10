
using Creo_TC_Live_Integration.TeamCenter;
using CreoToTc.Utils;
using Log;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Teamcenter.Services.Strong.Administration;
using Teamcenter.Services.Strong.Core;
using Teamcenter.Soa.Client.Model.Strong;

namespace AddToTc.CDAT_BulkUploader
{
    class Tc_Services
    {
        public static Teamcenter.Soa.Client.Connection tc_Connection = null;
        public static SessionService ss = null;
        public static Teamcenter.Services.Strong.Query.SavedQueryService savedQryServices = null;
        public static DataManagementService dmService = null;
        public static Teamcenter.Services.Strong.Query._2006_03.SavedQuery.GetSavedQueriesResponse savedQueries = null;
        public static ImanQuery uid_Qry = null;
        public static ImanQuery LatestItemRevQry = null;
        public static ImanQuery ItemRevisionQry = null;
        public static ImanQuery qry_3D_2D = null;
        public static Teamcenter.Services.Strong.Core.ReservationService reservService = null;
        public static SessionService session;
        public static Teamcenter.Services.Strong.Cad._2007_01.StructureManagement.StructureManagement sms2 = null;
        public static Teamcenter.Services.Strong.Query._2010_09.SavedQuery.SavedQuery queryService = null;
        public static PreferenceManagementService pmService = null;
        public static Teamcenter.Services.Strong.Query.FinderService fService = null;
        public static Teamcenter.Services.Strong.Bom.StructureManagementService bomSMService = null;
        public static Teamcenter.Services.Strong.Cad.DataManagementService dmCadService = null;
        public static Teamcenter.Services.Strong.Cad.StructureManagementService cadSMService = null;
        public static Teamcenter.Soa.Client.FileManagementUtility fileManag_Utility = null;

        public Tc_Services()
        {
            try
            {
                tc_Connection = Teamcenter.ClientX.Session.getConnection();

                ss = SessionService.getService(tc_Connection);

                savedQryServices = Teamcenter.Services.Strong.Query.SavedQueryService.getService(tc_Connection);

                dmService = DataManagementService.getService(tc_Connection);

                savedQueries = savedQryServices.GetSavedQueries();


                reservService = Teamcenter.Services.Strong.Core.ReservationService.getService(tc_Connection);
               
                session = SessionService.getService(tc_Connection);
                sms2 = Teamcenter.Services.Strong.Cad.StructureManagementService.getService(tc_Connection);
                queryService = Teamcenter.Services.Strong.Query.SavedQueryService.getService(tc_Connection);
                pmService = PreferenceManagementService.getService(tc_Connection);
                fService = Teamcenter.Services.Strong.Query.FinderService.getService(tc_Connection);
                bomSMService = Teamcenter.Services.Strong.Bom.StructureManagementService.getService(tc_Connection);
                dmCadService = Teamcenter.Services.Strong.Cad.DataManagementService.getService(tc_Connection);
                cadSMService = Teamcenter.Services.Strong.Cad.StructureManagementService.getService(tc_Connection);


                fileManag_Utility = new Teamcenter.Soa.Client.FileManagementUtility(Teamcenter.ClientX.Session.getConnection());
                dmService = DataManagementService.getService(Teamcenter.ClientX.Session.getConnection());
                Constants.ss = Teamcenter.Services.Strong.Structuremanagement.StructureService.getService(Teamcenter.ClientX.Session.getConnection());
            }
            catch (Exception ex)
            {
                log.write(logType.ERROR, "Upload exit in item creation function.");
                log.write(logType.ERROR, ex.ToString());
                Constants.globalError.Add("Teamcenter service creation failed");
            }
        }

    }
}
