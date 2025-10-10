
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Teamcenter.Services.Strong.Core;
using Teamcenter.Soa.Client.Model;
using Teamcenter.Soa.Client.Model.Strong;

namespace CreoToTc.Utils
{
    class Constants
    {

        public static string properties_ConfigFile = "Properties.config";
        //public static string bom_ConfigFile = "bom.config";
        public static List<String> globalError = new List<string>();       
        public static string inputLocation = "";
 
        public static String LOGIN = "LOGIN";
        public static String MAIN = "MAIN";
        public static String DE_STRING = "3DE_CREO: ";

        public static String TC_SERVER_HOST = "TC_SERVER_HOST";
        public static String TC_FSC_HOST = "";

        public static ImanQuery ItemRevisionQry = null;
      
        public static Teamcenter.Services.Strong.Structuremanagement.StructureService ss = null;
       // internal static string art_ConfigFile;
    }
}
