using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SolidEdge.StructureEditor.Interop;


namespace DemoAddInTC.se
{
    class SEECStuctureEditor
    {
        static String user = loginFromSE.userName;
        static String pwd = loginFromSE.password;
        static String URL = loginFromSE.URL;

        public static void perform_clone(String bstrItemID, String bstrItemRevID,String logFilePath)
        {

            SolidEdge.StructureEditor.Interop.Application StructureEditorApplication;
            StructureEditorApplication = (SolidEdge.StructureEditor.Interop.Application)Activator.CreateInstance(Type.GetTypeFromProgID("StructureEditor.Application"), true);
            if (StructureEditorApplication == null) return;
             SEECStructureEditorATP atp =  StructureEditorApplication.SEECStructureEditorATP;
            

            SEECStructureEditor SEECStructure = StructureEditorApplication.SEECStructureEditor;
            if (SEECStructure == null) return;

            utils.Utlity.Log("SEECStuctureEditor: Logging In to Teamcenter: ",logFilePath);
            utils.Utlity.Log("SEECStuctureEditor: Logging In to Teamcenter Group : " + loginFromSE.group, logFilePath);
            utils.Utlity.Log("SEECStuctureEditor: Logging In to Teamcenter Role : " + loginFromSE.role, logFilePath);
            int iret = SEECStructure.ValidateLogin(user, pwd, loginFromSE.group, loginFromSE.role, URL);
            utils.Utlity.Log("iRet: " + iret,logFilePath);
            //String bstrItemID = "6099553";
            //String bstrItemRevID = "04";
            String bstrFileName = bstrItemID + ".asm";
            String bstrRevisionRule = "Latest Working";
            String bstrFolderName = "";

            SEECStructure.Open(bstrItemID, bstrItemRevID, bstrFileName, bstrRevisionRule, bstrFolderName);
          
            SEECStructure.SetSaveAsAll();
            SEECStructure.AssignAll();
            SEECStructure.SetDataIntoAllCells("item_id", "000150");

            //SEECStructure.Close();
            SEECStructure.PerformActions();
            //SEECStructure.ClearCache();
            SEECStructure.Close();

            StructureEditorApplication.Quit();
            
        }

    }
}
