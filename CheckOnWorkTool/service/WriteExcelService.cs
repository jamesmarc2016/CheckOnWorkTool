using CheckOnWork.Controller;
using CheckOnWorkTool.domain;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.IO;

namespace CheckOnWorkTool.service
{
    class WriteExcelService
    {
        public Hashtable rowTables = new Hashtable();
        public Hashtable colTables = new Hashtable();

        //旷工信息
        public void writeStatusInfo(ArrayList statusMsgList, String path)
        {
            IWorkbook workbook = null;  //新建IWorkbook对象  
            FileStream fileStream = null;
            FileStream fileOut = null;
            try
            {
                fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);
                workbook = WorkbookFactory.Create(fileStream);
                ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表  
                //获取人所在行
                Hashtable nameRowTable = getNameRow(path);
                IRow numRow = sheet.GetRow(2);

                ICellStyle normalStyle = getNormalStyle(workbook);

                foreach (Object mlrs in statusMsgList)
                {
                    if (null != mlrs && mlrs is MonthlyLateRecord)
                    {
                        MonthlyLateRecord mlrd = (MonthlyLateRecord)mlrs;

                        String name = mlrd.Name;
                        String dep = mlrd.Dep;
                        String dateStr = mlrd.Date;
                        String msg = mlrd.Msg;
                        String clockType = mlrd.ClockType;

                        String nameDep = name + "-" + dep;
                        if (nameRowTable.Contains(nameDep))
                        {
                            int rowNum = (int)nameRowTable[nameDep];
                            //人在第几行
                            IRow row = sheet.GetRow(rowNum);
                            //获取日期列
                            IRow dateRow = sheet.GetRow(2);

                            int cellNum = 0;
                            for (int i = 0; i < dateRow.LastCellNum; i++)
                            {
                                ICell dateCell = dateRow.GetCell(i);
                                String dateCellStr = dateCell.ToString();
                                if (!"".Equals(dateCellStr) && dateStr.Equals(dateCellStr))
                                {
                                    cellNum = i;
                                    break;
                                }
                            }

                            if (0 == cellNum)
                            {
                                return;
                            }
                            else
                            {
                                ICell cell1 = sheet.GetRow(rowNum).CreateCell(cellNum);
                                ICell cell2 = sheet.GetRow(rowNum).CreateCell(cellNum + 1);

                                if (Properties.Resources.normal.Equals(clockType))
                                {
                                    cell1.SetCellValue("√");
                                    cell2.SetCellValue("√");

                                    cell1.CellStyle = normalStyle;
                                    cell2.CellStyle = normalStyle;
                                }
                                if (Properties.Resources.absent.Equals(clockType))
                                {
                                    

                                    cell1.SetCellValue("△");
                                    cell2.SetCellValue("△");

                                    cell1.CellStyle = normalStyle;
                                    cell2.CellStyle = normalStyle;
                                }
                                if (Properties.Resources.annualLeave.Equals(clockType))
                                {
                                    cell1.SetCellValue("年");
                                    cell2.SetCellValue("年");

                                    cell1.CellStyle = normalStyle;
                                    cell2.CellStyle = normalStyle;
                                }
                                if (Properties.Resources.waiqin.Equals(clockType) 
                                    || Properties.Resources.delay.Equals(clockType) 
                                    || Properties.Resources.checkOnWork.Equals(clockType)
                                    || Properties.Resources.leaveEarly.Equals(clockType))
                                {
                                    if (Properties.Resources.normal.Equals(msg))
                                    {
                                        cell1.SetCellValue("√");
                                        cell2.SetCellValue("√");

                                        cell1.CellStyle = normalStyle;
                                        cell2.CellStyle = normalStyle;
                                    }
                                    else
                                    {
                                        cell1.SetCellValue(msg);
                                        cell2.SetCellValue("√");
                                        
                                        cell1.CellStyle = normalStyle;
                                        cell2.CellStyle = normalStyle;
                                    }
                                }
                            }
                        }
                    }
                }
                fileOut = new FileStream(path, FileMode.Create);
                workbook.Write(fileOut);
            }
            catch (Exception ex)
            {
                Logging.Error(ex.Message + "-----" + ex.StackTrace);
                throw ex;
            }
            finally
            {
                this.closeStream(fileStream, fileOut, workbook);
            }
        }

        //常用格式
        private ICellStyle getNormalStyle(IWorkbook workbook)
        {
            ICellStyle style = workbook.CreateCellStyle();
            IFont font = workbook.CreateFont();
            //字体
            font.FontName = "宋体";
            //默认颜色
            font.Color = IndexedColors.Black.Index;
            //常规
            font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Normal;
            //12号
            font.FontHeightInPoints = 12;
            //背景色
            //style.FillPattern = FillPattern.SolidForeground;
            //style.FillForegroundColor = IndexedColors.LightGreen.Index;
            //// 居中
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            //边框
            style.BottomBorderColor = IndexedColors.Black.Index;
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style.SetFont(font);
            return style;
        }
        
        //获取姓名所在单元格
        private ICell getNameCol(ISheet sheet)
        {
            ICell cell = null;
            //循环行
            for (int i = 0;i < sheet.LastRowNum;i++)
            {
                IRow row = sheet.GetRow(i);
                //循环行
                for (int j = 0;j < row.LastCellNum;j++)
                {
                    cell = row.GetCell(j);
                    if (cell.ToString().Contains("姓名"))
                    {
                        goto END;
                    }
                }
            }
            END:;
            return cell;
        }

        //获取部门所在单元格
        private ICell getDepCol(ISheet sheet)
        {
            ICell cell = null;
            //循环行
            for (int i = 0; i < sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                //循环行
                for (int j = 0; j < row.LastCellNum; j++)
                {
                    cell = row.GetCell(j);
                    if (cell.ToString().Contains("部门"))
                    {
                        goto END;
                    }
                }
            }
            END:;
            return cell;
        }

        //获取部门-名字所在行
        private Hashtable getNameRow(string path)
        {
            Hashtable nameRowTables = new Hashtable();
            if (!"".Equals(path))
            {
                IWorkbook workbook = null;  //新建IWorkbook对象  
                FileStream fileStream = null;
                try
                {
                    fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);
                    workbook = WorkbookFactory.Create(fileStream);

                    ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表  

                    //获取姓名所在单元格
                    ICell nameCell = getNameCol(sheet);
                    //获取部门所在单元格
                    ICell depCell = getDepCol(sheet);

                    if (nameCell == null || depCell == null)
                    {
                        Logging.Error("未在写出文件找到部门或者姓名所在行。");
                        throw new Exception("未在写出文件找到部门或者姓名所在行。");
                    }

                    String depValue = "";
                    //循环获取行，排除无用信息
                    for (int i = nameCell.RowIndex + 1; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);

                        String depTemp = row.GetCell(depCell.ColumnIndex).ToString();
                        if (!"".Equals(depTemp))
                        {
                            depValue = depTemp;
                        }
                        String nameValue = row.GetCell(nameCell.ColumnIndex).ToString();

                        if ("".Equals(nameValue))
                        {
                            continue;
                        }
                        //名字table
                        nameRowTables.Add(nameValue + "-" + depValue, i);
                    }
                }
                catch (Exception ex)
                {
                    Logging.Error(ex.Message + "-----" + ex.StackTrace);
                    throw ex;
                }
                finally
                {
                    this.closeStream(fileStream,null,workbook);
                }
            }
            return nameRowTables;
        }

        //关闭流
        private void closeStream(FileStream fileIn, FileStream fileOut, IWorkbook workbook)
        {
            if (null != fileIn)
            {
                fileIn.Close();
            }
            if (null != fileOut)
            {
                fileOut.Close();
            }
            if (null != workbook)
            {
                workbook.Close();
            }
        }
    }
}
