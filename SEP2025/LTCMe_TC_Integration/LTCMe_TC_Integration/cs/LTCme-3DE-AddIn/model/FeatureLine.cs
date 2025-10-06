using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DemoAddInTC.model
{
    public class FeatureLine
    {
        public String EdgeBarName { get; set; }
        public String SystemName { get; set; }
        public String FeatureName { get; set; }

        [System.ComponentModel.DisplayName("PartName")]
        public String PartName { get; set; }
        public String Formula { get; set; }
        public String IsFeatureEnabled { get; set; }
        [System.ComponentModel.DisplayName("SKYname")]
        public String SKYname { get; set; }
        //{
        //    //get
        //    //{
        //    //    return this.IsFeatureEnabled;
        //    //}
        //    //set
        //    //{
        //    //    this.IsFeatureEnabled = value;
        //    //    if (this.IsFeatureEnabled != null && this.IsFeatureEnabled.Equals("") == false)
        //    //    {
        //    //        if (this.IsFeatureEnabled.Equals("Y", StringComparison.OrdinalIgnoreCase) == true)
        //    //        {
        //    //            this._S = "N";
        //    //        }
        //    //        else
        //    //        {
        //    //            this._S = "Y";
        //    //        }

        //    //    }
        //    //}
        //}
        //public String _S;
        
        public String SuppressionEnabled
        {
            get
            {
                if (this.IsFeatureEnabled != null && this.IsFeatureEnabled.Equals("") == false)
                {
                    if (this.IsFeatureEnabled.Equals("Y", StringComparison.OrdinalIgnoreCase) == true)
                    {
                       
                        return "N";
                    }
                    else
                    {                    
                        return "Y";
                    }
                }
                return "";
               
            }
            set
            {                
                // 28 - OCT, set is not working. No where this property is Set.
                if (this.SuppressionEnabled != null && this.SuppressionEnabled.Equals("") == false)
                {
                    if (this.SuppressionEnabled.Equals("Y", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        IsFeatureEnabled = "N";
                    }
                    else
                    {
                        IsFeatureEnabled = "Y";
                    }
                }
                              
            }
        }

    }
}
