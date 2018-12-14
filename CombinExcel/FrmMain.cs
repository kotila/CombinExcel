using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.XSSF.UserModel;
using NPOI.Util;
using System.IO;
using NPOI.SS.UserModel;
using System.Diagnostics;
using System.Configuration;

namespace CombinExcel
{
    public partial class FrmMain : Form
    {
        string strFileType = "";
        string strFileName ;
        String thisPhoneNumber;
        public FrmMain()
        {
            InitializeComponent();
            textBox1.Text = "";
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx|所有文件|*.*";
                        ofd.ValidateNames = true;
            ofd.CheckPathExists = true;
            ofd.CheckFileExists = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                strFileName = ofd.FileName.ToLower();
                strFileType = Path.GetExtension(strFileName).ToLower();
                //其他代码
                textBox1.Text = strFileName;
            }
            if (!Utility.IsAllowedExtension(strFileName))
            {
                MessageBox.Show("错误的文件格式！当前版本只允许使用*.xlsx或*.xls格式文件");
            }
        }

        private void BtnCombinData_Click(object sender, EventArgs e)
        {
            //获取Configuration对象
            //Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            //open file:
            IWorkbook workbook = null;
            //ISheet sheet = null;
            IRow row1= null;
            IRow row2 = null;
            //ICell cell = null;

            //int startRow = 0;
            //读歌华列表中的业务号码是第几列
            int Gehua_PhoneColNumber = int.Parse(ConfigurationManager.AppSettings["Gehua_PhoneColNumber"]);
            int Unicom_PhoneColNumber = int.Parse(ConfigurationManager.AppSettings["Unicom_PhoneColNumber"]);

            int read1 = int.Parse(ConfigurationManager.AppSettings["ReadCol1"]);
            int read2 = int.Parse(ConfigurationManager.AppSettings["ReadCol2"]);
            int read3 = int.Parse(ConfigurationManager.AppSettings["ReadCol3"]);

            int write1 = int.Parse(ConfigurationManager.AppSettings["WriteCol1"]);
            int write2 = int.Parse(ConfigurationManager.AppSettings["WriteCol2"]);
            int write3 = int.Parse(ConfigurationManager.AppSettings["WriteCol3"]);

            FileStream fs;
            string filepath;
            if (textBox1.Text.Trim() != "")
            {
                //打开选择的Excel文件
                //根据现有的Excel文档创建工作簿
                filepath = textBox1.Text;
                fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
                if (strFileType==".xlsx")
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else
                {
                    workbook = new HSSFWorkbook(fs);
                }

                fs.Close();
                try
                {
                    //获得工作表
                    //string g_Sheet=  ConfigurationManager.AppSettings["Gehua_sheet"];
                    ISheet sheet1 = workbook.GetSheetAt(int.Parse(ConfigurationManager.AppSettings["Gehua_sheet"]));

                    //string u_Sheet = ConfigurationManager.AppSettings["Unicom_sheet"];
                    ISheet sheet2 = workbook.GetSheetAt(int.Parse(ConfigurationManager.AppSettings["Unicom_sheet"]));

                    //获得所有行的集合
                    System.Collections.IEnumerator rows1 = sheet1.GetRowEnumerator();
                    rows1.MoveNext();
                    
                    //int rowCount = 0;
                    while (rows1.MoveNext())
                    {
                        //rowCount = rowCount + 1;
                        if (strFileType == ".xlsx")
                        {
                            row1 = (XSSFRow)rows1.Current;
                        }
                        else
                        {
                            row1 = (HSSFRow)rows1.Current;
                        }
                        //自动获取业务号码所在列
                        if (Gehua_PhoneColNumber < 0)
                        {
                            Gehua_PhoneColNumber = GetTitleNumber(row1, "业务号码");
                        }

                        if (Gehua_PhoneColNumber>=0)
                        {
                            Boolean bFlag = true;

                            ICell cell = row1.GetCell(Gehua_PhoneColNumber);
                            if (cell is null)
                            {
                                break;
                            }
                            cell.SetCellType(CellType.String);
                            thisPhoneNumber = cell.StringCellValue;

                            //判断读取的业务号码是否为11位的电话号码:
                            int phoneLength = thisPhoneNumber.Length;
                            if (phoneLength > 10 && phoneLength < 13)
                            {
                                thisPhoneNumber = thisPhoneNumber.Substring(0, 11);
                                
                                bFlag = ulong.TryParse(thisPhoneNumber, out ulong num);
                            }
                            else
                            {
                                bFlag = false;
                            }
                            //是电话号码
                            if (bFlag)
                            {
                                System.Collections.IEnumerator rows2 = sheet2.GetRowEnumerator();
                                rows2.MoveNext();
                                while (rows2.MoveNext())
                                {
                                    if (strFileType == ".xlsx")
                                    {
                                        row2 = (XSSFRow)rows2.Current;
                                    }
                                    else
                                    {
                                        row2 = (HSSFRow)rows2.Current;
                                    }
                                    row2.GetCell(Unicom_PhoneColNumber).SetCellType(CellType.String);
                                    string myPhone = row2.GetCell(Unicom_PhoneColNumber).StringCellValue;
                                    if (myPhone.Length < 11)
                                    {
                                        break;
                                    }
                                    if (thisPhoneNumber.Trim().Substring(0, 11) == myPhone.Trim().Substring(0, 11))
                                    {
                                        row1.GetCell(write1).SetCellValue(row2.GetCell(read1).NumericCellValue);
                                        row1.GetCell(write2).SetCellValue(row2.GetCell(read2).NumericCellValue);
                                        row1.GetCell(write3).SetCellValue(row2.GetCell(read3).NumericCellValue);
                                        //workbook.Write(fs);
                                        break;
                                    }
                                }
                            }
                        }

                    }
                }
                catch (Exception)
                {
                    fs.Close();
                    throw;
                }

            }
            else
            {
                MessageBox.Show("请先选择要处理的文件！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            string temp = DateTime.Now.ToUniversalTime().ToString();
                temp=temp .Replace("-","").Replace(":","").Replace("/","").Replace(" ","");   // 2008-9-4 12:19:14
            string fileHead = ConfigurationManager.AppSettings["FileHead"];
            filepath = fileHead + temp + strFileType;

            fs = new FileStream(filepath, FileMode.Create, FileAccess.Write);
            workbook.Write(fs);
            fs.Close();
            MessageBox.Show("文件已保存到：" + filepath);
        }

        private static int GetTitleNumber(IRow row, string keyWord)
        {
            int cellNo=-1;
            for (int i = 0; i < row.Cells.Count; i++)
            {
                if (null != row.GetCell(i))
                {
                    ICell cell = row.GetCell(i);
                    cell.SetCellType(CellType.String);
                    string tmp = cell.StringCellValue;
                    if (tmp == keyWord.Trim())
                    {
                        cellNo = i;
                        break;
                    }
                }
            }

            return cellNo;
        }
    }
}
