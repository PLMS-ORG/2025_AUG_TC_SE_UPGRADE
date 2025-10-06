using DemoAddInTC.se;
using Microsoft.VisualBasic.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Teamcenter.Soa.Client.Model.Strong;

namespace DemoAddInTC.utils
{
    internal class Tc_Process
    {

        public static ImanQuery getTcQuery(String queryToFind,string logFilepath)
        {
            try
            {
                ImanQuery tcQueryToReturn = null;


                if (TcAdaptor.savedQryServices == null)
                {
                    Utility.Log("Error:TcAdaptor.savedQryServices is null", logFilepath);
                    return null;
                }


                Teamcenter.Services.Strong.Query._2006_03.SavedQuery.GetSavedQueriesResponse savedQueries = TcAdaptor.savedQryServices.GetSavedQueries();

                if (savedQueries.Queries.Length == 0)
                {
                    Utility.Log( "Failed to get saved queries", logFilepath);
                    return null;
                }

                else
                {
                    for (int i = 0; i < savedQueries.Queries.Length; i++)
                    {
                        if (savedQueries.Queries[i].Name.Equals(queryToFind))
                        {
                            tcQueryToReturn = savedQueries.Queries[i].Query;

                            break;
                        }
                    }
                }
                return tcQueryToReturn;
            }
            catch (Exception ex)
            {
                Utility.Log("Exception in getTcQuery :" + ex, logFilepath);
                return null;
            }
        }
    }
}
