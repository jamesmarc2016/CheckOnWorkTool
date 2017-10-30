using CheckOnWork.Controller;
using CheckOnWorkTool.domain;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace CheckOnWorkTool.service
{
    class MonthlyService
    {
        public ArrayList getStatusMsg(String path)
        {
            ArrayList statusMsgList = new ArrayList();

            if (!"".Equals(path))
            {
                IWorkbook workbook = null;  //新建IWorkbook对象  
                FileStream fileStream = null;
                try
                {
                    fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);
                    workbook = WorkbookFactory.Create(path);

                    ISheet sheet = workbook.GetSheetAt(0);  //获取第一个考勤记录表  

                    //获取天数所在行
                    Hashtable dateTable = new Hashtable();
                    IRow dateRow = sheet.GetRow(2);
                    for (int i = 0; i < dateRow.LastCellNum; i++)
                    {
                        String cellStr = dateRow.GetCell(i).ToString();
                        if (cellStr.Contains("日"))
                        {
                            String date = cellStr.Substring(0, cellStr.IndexOf("日"));
                            //date-cellNum
                            dateTable.Add(date, i);
                        }
                    }

                    //循环获取行，排除无用信息
                    for (int i = 5; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);

                        if (null != row.GetCell(0) && null != row.GetCell(1))
                        {
                            String name = row.GetCell(0).ToString();
                            String dep = row.GetCell(1).ToString();

                            foreach (String dateStr in dateTable.Keys)
                            {
                                MonthlyLateRecord mlinfo = new MonthlyLateRecord();
                                mlinfo.Name = name;
                                mlinfo.Dep = dep;
                                //日期
                                mlinfo.Date = dateStr;

                                int cellNum = (int)dateTable[dateStr];
                                ICell cell = row.GetCell(cellNum);

                                if (null == cell)
                                {
                                    continue;
                                }

                                //批注
                                IComment comment = cell.CellComment;
                                if (null == comment)
                                {
                                    continue;
                                }

                                String cellStr = cell.ToString();

                                //正常
                                if (cellStr.Contains("√"))
                                {
                                    //打卡类型
                                    mlinfo.ClockType = Properties.Resources.normal;
                                    mlinfo.Msg = getNormalMsg(comment);
                                    statusMsgList.Add(mlinfo);
                                    continue;
                                }
                                //旷工信息
                                if (cellStr.Contains(Properties.Resources.absent))
                                {
                                    mlinfo.ClockType = Properties.Resources.absent;
                                    mlinfo.Msg = Properties.Resources.absent;
                                    statusMsgList.Add(mlinfo);
                                    continue;
                                }
                                //外勤
                                if (cellStr.Contains(Properties.Resources.waiqin))
                                {
                                    mlinfo.ClockType = Properties.Resources.waiqin;
                                    mlinfo.Msg = getWaiQinMsg(comment);
                                    statusMsgList.Add(mlinfo);
                                    continue;
                                }
                                //迟到
                                if (cellStr.Contains(Properties.Resources.delay))
                                {
                                    mlinfo.ClockType = Properties.Resources.delay;
                                    mlinfo.Msg = getDelayMsg(comment);
                                    statusMsgList.Add(mlinfo);
                                    continue;
                                }
                                //考勤
                                if (cellStr.Contains(Properties.Resources.checkOnWork))
                                {
                                    mlinfo.ClockType = Properties.Resources.checkOnWork;
                                    mlinfo.Msg = getCheckOnWork(comment);
                                    statusMsgList.Add(mlinfo);
                                    continue;
                                }
                                //早退
                                if (cellStr.Contains(Properties.Resources.leaveEarly))
                                {
                                    mlinfo.ClockType = Properties.Resources.leaveEarly;
                                    mlinfo.Msg = getCheckOnWork(comment);
                                    statusMsgList.Add(mlinfo);
                                    continue;
                                }
                                //年休假
                                if (cellStr.Contains(Properties.Resources.annualLeave))
                                {
                                    mlinfo.ClockType = Properties.Resources.annualLeave;
                                    mlinfo.Msg = Properties.Resources.annualLeave;
                                    statusMsgList.Add(mlinfo);
                                    continue;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logging.Error(ex.Message + "-----" + ex.StackTrace);
                    throw ex;
                }
                finally
                {
                    if (null != fileStream)
                    {
                        fileStream.Close();
                    }
                    if (null != workbook)
                    {
                        workbook.Close();
                    }
                }
            }
            return statusMsgList;
        }
        
        //获取签到信息
        private String getSignMsg(String[] msgArr,Regex reg)
        {
            ArrayList timeList = new ArrayList();
            foreach (String msg in msgArr)
            {
                Match match = reg.Match(msg);
                if (!reg.IsMatch(msg))
                {
                    continue;
                }

                String value = match.Value;
                //解析时间
                DateTime clockTime = DateTime.Parse(value);

                if (timeList.Count == 0)
                {
                    timeList.Add(clockTime);
                }
                else
                {
                    DateTime temp = (DateTime)timeList[0];
                    if (clockTime.CompareTo(temp) < 0)
                    {
                        timeList.RemoveAt(0);
                        timeList.Insert(0, clockTime);
                    }
                }
            }
            if (timeList.Count == 0)
            {
                return null;
            }
            return timeDelay((DateTime)timeList[0]);
        }

        //考勤
        private String getLeaveEarly(IComment comment)
        {
            String commentStr = comment.String.ToString();
            String[] msgArr = Regex.Split(commentStr, "\n", RegexOptions.IgnoreCase);
            Regex reg = new Regex(@"[0-9][0-9]:[0-5][0-9]");
            return getSignMsg(msgArr, reg);
        }

        //考勤
        private String getCheckOnWork(IComment comment)
        {
            String commentStr = comment.String.ToString();
            String[] msgArr = Regex.Split(commentStr, "\n", RegexOptions.IgnoreCase);
            Regex reg = new Regex(@"[0-9][0-9]:[0-5][0-9]");
            return getSignMsg(msgArr, reg);
        }

        //正常打卡，检查错误
        private String getNormalMsg(IComment comment)
        {
            String commentStr = comment.String.ToString();
            String[] msgArr = Regex.Split(commentStr, "\n", RegexOptions.IgnoreCase);
            Regex reg = new Regex(@"[0-9][0-9]:[0-5][0-9]");
            return getSignMsg(msgArr, reg);
        }

        //迟到
        private String getDelayMsg(IComment comment)
        {
            String commentStr = comment.String.ToString();
            String[] msgArr = Regex.Split(commentStr, "\n", RegexOptions.IgnoreCase);
            Regex reg = new Regex(@"[0-9][0-9]:[0-5][0-9]");
            return getSignMsg(msgArr, reg);
        }

        //外勤
        private String getWaiQinMsg(IComment comment)
        {
            String commentStr = comment.String.ToString();
            String[] msgArr = Regex.Split(commentStr, "\n", RegexOptions.IgnoreCase);

            Regex reg = new Regex(@"[0-9][0-9][0-9][0-9]-[0-1][0-9]-[0-3][0-9] [0-1]?[0-9]:[0-5][0-9]");

            String result = getSignMsg(msgArr, reg);

            //判断迟到是否超过30分钟
            if (null != result && result.Contains(Properties.Resources.delay))
            {
                String temp = result.Replace(Properties.Resources.delay, "").Replace("分钟", "").Trim();
                if (int.Parse(temp) > 30)
                {
                    foreach (String msg in msgArr)
                    {
                        //签到且不是未打卡
                        if (msg.Contains(Properties.Resources.sign))
                        {
                            if (!msg.Contains(Properties.Resources.noClock))
                            {
                                reg = new Regex(@"[0-9][0-9]:[0-5][0-9]");
                                result = getSignMsg(msgArr, reg);
                            }
                        }
                    }
                }
            }
            return result;
        }

        //解析时间
        private String timeDelay(DateTime clockTime)
        {
            StringBuilder build = new StringBuilder();

            String hourStr = clockTime.ToString("HH");
            String minutStr = clockTime.ToString("mm");
            int hour = int.Parse(hourStr);
            int minut = int.Parse(minutStr);

            if ((hour == 9 && minut > 0) || hour > 9)
            {
                int delayTime = (hour - 9) * 60 + minut;
                build.Append(Properties.Resources.delay);
                build.Append(delayTime);
                build.Append("分钟");
            }
            else
            {
                build.Append(Properties.Resources.normal);
            }
            return build.ToString();
        }
    }
}
