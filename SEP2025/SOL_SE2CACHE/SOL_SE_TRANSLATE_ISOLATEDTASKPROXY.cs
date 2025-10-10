using SolidEdgeCommunity.Extensions;
using SolidEdgeCommunity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LTC_SE2CACHE 
{
    class SOL_SE_TRANSLATE_ISOLATEDTASKPROXY : IsolatedTaskProxy
    {

        //struct SE_TRANSLATE_OPTIONS
        //{
        //    public String StageDir;
        //    public String InputFile;
        //    public String Format1;
        //    public String logFilePath;
        //    public String OutputFolder;
        //    public String Format2;

        //}

        /* In case of future - RPC Server not Available Issues - Changes to be DONE ----
        // https://github.com/SolidEdgeCommunity/SolidEdge.Community
         * https://github.com/SolidEdgeCommunity/Samples/blob/master/General/OpenSave/cs/OpenSave/OpenSaveTask.cs
         * https://github.com/SolidEdgeCommunity/SolidEdge.Community/blob/master/src/SolidEdge.Community/IsolatedTaskProxy.cs
         * 
        */

        //public bool SaveDraftAsIsolatedTask(String StageDir, String InputFile, String Format, String logFilePath, String OutputFolder,SE_TRANSLATE_OPTIONS options)
        //{
        //    //SE_TRANSLATE_OPTIONS options;
        //    options.StageDir = StageDir;
        //    options.InputFile = InputFile;
        //    options.Format1 = Format;
        //    options.logFilePath = logFilePath;
        //    options.OutputFolder = OutputFolder;

        //    InvokeSTAThread<SE_TRANSLATE_OPTIONS>(SaveDraftAsWithOptions, options);
        //    return true;
        //}

        //public static void SaveDraftAsWithOptions(SE_TRANSLATE_OPTIONS options)
        //{
        //    return ;
        //}

    }
}
