using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DemoAddInTC.utils
{
    class StringUtils
    {
        public static String RemoveUnitsFromDimension(String Dimension)
        {
            char[] spaceSeparator = new char[] { ' ' };
            if (Dimension != null && Dimension.Equals("") == false)
            {
                String[] DimensionArr = Dimension.Split(spaceSeparator);
                if (DimensionArr != null && DimensionArr.Length > 0) {
                    return DimensionArr[0];
                }else {
                    return Dimension;
                }
            }

            return Dimension;
        }
    }
}
