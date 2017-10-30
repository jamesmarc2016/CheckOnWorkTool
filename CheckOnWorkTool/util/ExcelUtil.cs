using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CheckOnWorkTool.util
{
    class ExcelUtil
    {
        public static bool isNumber(string str)
        {
            Regex reg = new Regex(@"^\d+([.]\d+)?$");
            return reg.Match(str).Success;
        }

        
    }
}
