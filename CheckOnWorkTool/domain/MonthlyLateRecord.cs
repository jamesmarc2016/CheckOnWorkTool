using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CheckOnWorkTool.domain
{
    class MonthlyLateRecord
    {
        //员工名字
        private String name;
        //部门
        private String dep;
        //日期
        private String date;
        //状态
        private String status;
        //打卡类型
        private String clockType;
        //附加消息
        private String msg;

        public string Name
        {
            get
            {
                return name;
            }

            set
            {
                name = value;
            }
        }

        public string Dep
        {
            get
            {
                return dep;
            }

            set
            {
                dep = value;
            }
        }

        public string Date
        {
            get
            {
                return date;
            }

            set
            {
                value = value.Substring(value.LastIndexOf("-") + 1, value.Length - value.LastIndexOf("-") - 1);
                if (value.StartsWith("0"))
                {
                    value = value.Substring(value.LastIndexOf("0") + 1, value.Length - value.LastIndexOf("0") - 1);
                }
                date = value;
            }
        }

        public string Status
        {
            get
            {
                return status;
            }

            set
            {
                status = value;
            }
        }

        public string ClockType
        {
            get
            {
                return clockType;
            }

            set
            {
                clockType = value;
            }
        }

        public string Msg
        {
            get
            {
                return msg;
            }

            set
            {
                msg = value;
            }
        }
    }
}
