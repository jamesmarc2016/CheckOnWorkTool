using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CheckOnWorkTool.util
{
    public class MessagesUtil
    {
        public static void addMsg(String msg)
        {
            if (null == main.textBox4.Text || "".Equals(main.textBox4.Text))
            {
                main.textBox4.Text += msg;
            }else
            {
                main.textBox4.Text += "\r\n" + msg;
            }
            
        }
    }
}
