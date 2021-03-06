﻿using Aras.IOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace manhour_report
{
    public partial class 工时统计 : Form
    {
        public 工时统计()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //设置dateTimePicker的数据格式
            dateTimePicker_report.Format = DateTimePickerFormat.Custom;
            dateTimePicker_report.CustomFormat = "yyyy年MM月";
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            button1.Enabled = false;
            textBox_message.Text = "报表生成中.......";
            dateTimePicker_report.Focus();
            //获取输入的时间
            DateTime dt = dateTimePicker_report.Value;

            //用来限定访问数据的范围
            string firstDay = dt.Year + "-" + dt.Month + "-" + 1;
            string lastDay;
            if (dt.Month.Equals(12))
            {
                lastDay = (dt.Year + 1) + "-" + 1 + "-" + 1;
            }
            else
            {
                lastDay = dt.Year + "-" + (dt.Month + 1) + "-" + 1;
            }
            try
            {
                // 访问Aras
                string url = Properties.Settings.Default.url;
                string username = Properties.Settings.Default.username;
                string pwd = Properties.Settings.Default.pwd;
                string database = Properties.Settings.Default.database;

                Item loginItem = IomFactory.CreateHttpServerConnection(url, database, username, pwd).Login();
                if (loginItem.isError())
                {
                    MessageBox.Show("访问数据库失败,请重试");
                    button1.Enabled = true;
                    textBox_message.Text = String.Empty;
                }
                else
                {
                    Innovator inn = loginItem.getInnovator();

                    //EB25E105EFDF478980005105489F6E74表示项目List的cofigId;
                    Item projectSql = inn.applySQL("SELECT value FROM[IPD].[innovator].VALUE where SOURCE_ID = 'EB25E105EFDF478980005105489F6E74' ");

                    //D1593C4073EE49218E67A6B9F28CA943表示用户list的configId
                    Item nameSql = inn.applySQL("SELECT value,label  FROM [IPD].[innovator].VALUE a where SOURCE_ID = 'D1593C4073EE49218E67A6B9F28CA943' ");
                    Dictionary<string, string> nameMap = new Dictionary<string, string>();
                    for (int i = 0; i < nameSql.getItemCount(); i++)
                    {
                        Item item = nameSql.getItemByIndex(i);
                        string nameE = item.getProperty("value");
                        string nameC = item.getProperty("label");
                        if (nameE != null)
                        {
                            nameMap[nameE] = nameC;
                        }
                    }

                    Item allDateSql = inn.applySQL("SELECT a.[SOURCE_ID] " +
                                                  ", a.[PROJECT_NAME]  " +
                                                  ", a.[PROJECT_TIME] " +
                                                  ", a.[PROJECT_COMMENT] " +
                                                  ", b.MY_NAME " +
                                                  " , b.WORKING_DATE " +
                                                  "FROM[IPD].[innovator].[LAUREL_PROJECT_AND_TIME] a " +
                                                  "Left Join IPD.innovator.MANHOUR_REGISTER b on b.CONFIG_ID = a.SOURCE_ID " +
                                                  "and b.WORKING_DATE > ' " + firstDay + " ' and b.WORKING_DATE <'" + lastDay + "' ");

                    if (allDateSql.getItemCount() == -1 || allDateSql.getItemCount() == 0)
                    {
                        MessageBox.Show("未查询到数据,请重试");
                        button1.Enabled = true;
                        textBox_message.Text = String.Empty;
                        return;
                    }

                    //把所有的数据制作成Map
                    Dictionary<string, List<Item>> dictionary = new Dictionary<string, List<Item>>();
                    for (int i = 0; i < allDateSql.getItemCount(); i++)
                    {
                        Item item = allDateSql.getItemByIndex(i);
                        string projectName = item.getProperty("project_name");
                        string myName = item.getProperty("my_name");
                        if (!string.IsNullOrWhiteSpace(projectName) && !string.IsNullOrWhiteSpace(myName))
                        {
                            if (dictionary.ContainsKey(projectName))
                            {
                                List<Item> list = dictionary[projectName];
                                list.Add(item);
                            }
                            else
                            {
                                List<Item> list = new List<Item>();
                                list.Add(item);
                                dictionary[projectName] = list;
                            }
                        }
                    }
                    //新建Excel和工作簿
                    Microsoft.Office.Interop.Excel.Application oXL;
                    Microsoft.Office.Interop.Excel._Workbook oWB;

                    //启动Excel并获取应用程序对象
                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    oXL.Visible = false;
                    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                    //遍历所有的名字
                    for (int i = 0; i < projectSql.getItemCount(); i++)
                    {
                        Item item = projectSql.getItemByIndex(i);
                        string name = item.getProperty("value"); //map制作完成,
                                                                 //防止list中的项目在Map中不存在
                        if (dictionary.ContainsKey(name))
                        {
                            List<Item> sqlResult = dictionary[name];
                            if (sqlResult.Count() <= 0)
                            {
                                MessageBox.Show(name + "无查询结果");
                                button1.Enabled = true;
                                textBox_message.Text = String.Empty;
                            }
                            else
                            {
                                makeSheet(sqlResult, dt, oXL, oWB, name, nameMap);

                            }
                        }
                    }
                    oXL.Visible = true;
                    oXL.UserControl = true;
                    this.Close();

                }
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.InnerException);
                errorMessage = String.Concat(errorMessage, theException.Source);
                errorMessage = String.Concat(errorMessage, theException.ToString());
                errorMessage = String.Concat(errorMessage, theException.TargetSite);
                MessageBox.Show(errorMessage, "Error");
            }
        }

        private void makeSheet(List<Item> sqlResult, DateTime dt, Microsoft.Office.Interop.Excel.Application oXL, Microsoft.Office.Interop.Excel._Workbook oWB, string name, Dictionary<string, string> nameMap)
        {
            int dtCount = DateTime.DaysInMonth(dt.Year, dt.Month);
            //获取访问数据库的数据
            HashSet<string> nameSet = new HashSet<string>();
            for (int i = 0; i < sqlResult.Count(); i++)
            {
                nameSet.Add(sqlResult[i].getProperty("my_name"));
            }
            _Worksheet oSheet;
            try
            {
                //数据选项菜单
                oXL.Sheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                oSheet = (_Worksheet)oWB.ActiveSheet;
                oSheet.Name = name;

                //逐单元添加表标题
                oSheet.Cells[1, 1] = "日期 姓名";
                //纵坐标,日期
                for (int i = 0; i < dtCount; i++)
                {
                    oSheet.Cells[i + 2, 1] = dt.Year + "-" + dt.Month + "-" + (i + 1);
                }
                //横坐标,人名
                for (int i = 0; i < nameSet.Count; i++)
                {
                    if (nameMap.ContainsKey(nameSet.ElementAt(i)))
                    {
                        oSheet.Cells[1, i + 2] = nameMap[nameSet.ElementAt(i)];
                    }
                    else
                    {
                        //
                        oSheet.Cells[1, i + 2] = nameSet.ElementAt(i);
                    }
                }

                //遍历整个结果集,将结果填写到相应的位置
                for (int i = 0; i < sqlResult.Count(); i++)
                {
                    Item hoursItem = sqlResult[i];
                    string myName = hoursItem.getProperty("my_name");
                    string dateTimeStr = hoursItem.getProperty("working_date");
                    int y = -1;
                    //遍历所有的名字
                    for (int j = 0; j < nameSet.Count; j++)
                    {
                        if (myName.Equals(nameSet.ElementAt(j)))
                        {
                            y = j + 2;
                        }
                    }
                    if (y == -1)
                    {
                        MessageBox.Show("程序出错啦");
                        button1.Enabled = true;
                        textBox_message.Text = String.Empty;
                        return;
                    }
                    DateTime dateTime = DateTime.Parse(dateTimeStr);
                    int day = dateTime.Day;
                    //设置y
                    int x = day + 1;
                    string hour = hoursItem.getProperty("project_time");
                    oSheet.Cells[x, y] = hour;
                }
                //Format A1:D1 as bold, vertical alignment = center.
                //第一行
                Range rowRange = oSheet.get_Range("A1", "BZ1");
                rowRange.Font.Bold = true;
                rowRange.VerticalAlignment =
                    XlVAlign.xlVAlignCenter;
                rowRange.EntireColumn.AutoFit();
                
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.InnerException);
                errorMessage = String.Concat(errorMessage, theException.Source);
                errorMessage = String.Concat(errorMessage, theException.ToString());
                errorMessage = String.Concat(errorMessage, theException.TargetSite);
                MessageBox.Show(errorMessage, "Error");
            }
            
        }

        private void textBox_message_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
