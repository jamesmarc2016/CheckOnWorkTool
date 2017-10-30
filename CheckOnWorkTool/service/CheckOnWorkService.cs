using CheckOnWork.Controller;
using CheckOnWorkTool.util;
using System;
using System.Collections;
using System.Windows.Forms;

namespace CheckOnWorkTool.service
{
    class CheckOnWorkService
    {
        public static int currentThreadNum = 0;
        public void writeMegToExcel(String writeExcelFilePath, String monthlyFilePath)
        {
            if ("".Equals(writeExcelFilePath) || "".Equals(monthlyFilePath))
            {
                MessageBox.Show("文件路径不能为空");
                //执行完程序重置次数
                currentThreadNum = 0;
                return;
            }
            else if (currentThreadNum != 1)
            {
                MessageBox.Show("一次只能处理一个请求");
                //执行完程序重置次数
                currentThreadNum = 0;
                return;
            }
            else
            {
                try
                {
                    //获取月度考勤信息
                    MessagesUtil.addMsg("获取月度考勤信息...");
                    ArrayList statusMsgList =  new MonthlyService().getStatusMsg(monthlyFilePath);
                    MessagesUtil.addMsg("月度考勤信息获取成功.");
                    
                    //写出信息到写出文件
                    WriteExcelService wes = new WriteExcelService();
                    MessagesUtil.addMsg("写出信息到写出文件...");
                    wes.writeStatusInfo(statusMsgList,writeExcelFilePath);
                    MessagesUtil.addMsg("写出信息到写出文件成功.");
                }
                catch (Exception ex)
                {
                    Logging.Error(ex.Message + "-----" + ex.StackTrace);
                    MessageBox.Show("打开Excel失败! 失败原因：" + ex.Message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally
                {
                    //执行完程序重置次数
                    currentThreadNum = 0;
                }
            }
        }
    }
}
