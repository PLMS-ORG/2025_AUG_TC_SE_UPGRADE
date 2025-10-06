using SolidEdgeCommunity.AddIn;
using SolidEdgeCommunity.Extensions; // https://github.com/SolidEdgeCommunity/SolidEdge.Community/wiki/Using-Extension-Methods
using SolidEdgeFramework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace DemoAddInTC
{
    class Ribbon3d : SolidEdgeCommunity.AddIn.Ribbon
    {
        const string _embeddedResourceName = "DemoAddInTC.Ribbon3d.xml";
       

        public Ribbon3d()
        {
            // Get a reference to the current assembly. This is where the ribbon XML is embedded.
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();

            // In this example, XML file must have a build action of "Embedded Resource".
            this.LoadXml(assembly, _embeddedResourceName);            

            // Get the Solid Edge version.
            var version = DemoAddInTC.Instance.SolidEdgeVersion;

            // View.GetModelRange() is only available in ST6 or greater.
            if (version.Major < 106)
            {
                
            }
        }

        public override void OnControlClick(RibbonControl control)
        {            

            

            
        }

        

        private MyViewOverlay GetActiveOverlay()
        {
            var controlller = DemoAddInTC.Instance.ViewOverlayController;
            var window = (SolidEdgeFramework.Window)DemoAddInTC.Instance.Application.ActiveWindow;
            var overlay = (MyViewOverlay)controlller.GetOverlay(window);

            if (overlay == null)
            {
                // If the overlay has not been created yet, add a new one.
                overlay = controlller.Add<MyViewOverlay>(window);
            }

            return overlay;
        }
    }
}
