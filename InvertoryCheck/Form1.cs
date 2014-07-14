using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace InvertoryCheck
{
    public partial class Form1 : Form
    {
        private Dictionary<string, int> dictInventory;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dictInventory = new Dictionary<string, int>();
            lstshijikucun.Items.Clear();
            lblxitongkucun.Text = string.Empty;
        }

        private void btnxitongkucun_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult d = openFileDialog1.ShowDialog();
                if (d == DialogResult.OK)
                {
                    lblxitongkucun.Text = openFileDialog1.FileName;
                    SubXiTongKuCun(openFileDialog1.FileName);
                }
            }
            catch (Exception ex)
            {
                lblxitongkucun.Text = string.Empty;
                MessageBox.Show("导入excel出错，请检查数据(B列为款号，P列为库存)\n\r\n\r" + ex.Message);
            }
        }

        private void btnshijikucun_Click(object sender, EventArgs e)
        {

            try
            {
                DialogResult d = openFileDialog2.ShowDialog();
                if (d == DialogResult.OK)
                {
                    if (lstshijikucun.Items.Contains(openFileDialog2.FileName) == false)
                    {
                        lstshijikucun.Items.Add(openFileDialog2.FileName);
                        AddShiJiKuCun(openFileDialog2.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                lstshijikucun.Items.Remove(openFileDialog2.FileName);
                MessageBox.Show("导入excel出错，请检查数据(A列为款号，B列为库存)\n\r\n\r"+ex.Message);
            }
        }

        private void btnjisuankucun_Click(object sender, EventArgs e)
        {
            ExcelHelp ex = new ExcelHelp();
            try
            {
                //string localFilePath, fileNameExt, newFileName, FilePath; 
                SaveFileDialog sfd = new SaveFileDialog();
                //设置文件类型 
                sfd.Filter = "Excel Files（*.xls）|*.xls";
                //设置默认文件类型显示顺序 
                sfd.FilterIndex = 1;

                //保存对话框是否记忆上次打开的目录 
                sfd.RestoreDirectory = true;

                //点了保存按钮进入 
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string localFilePath = sfd.FileName.ToString(); //获得文件路径 
                    string fileNameExt = localFilePath.Substring(localFilePath.LastIndexOf("\\") + 1); //获取文件名，不带路径

                 
                    ex.Create();
                    ex.ws = (Excel.Worksheet)ex.wb.ActiveSheet;
                    int rowNumber = 1;
                    foreach (KeyValuePair<string, int> pair in dictInventory)
                    {
                        string key = pair.Key;
                        int value = pair.Value;
                        ex.SetCellValue(rowNumber, 1, key);
                        ex.SetCellValue(rowNumber, 2, value);
                        rowNumber++;
                    }
                    ex.SaveAs(localFilePath);
                    MessageBox.Show("已导出excel文件 " + localFilePath);
                }
            }
            catch (Exception e1)
            {
                MessageBox.Show("计算excel出错，请检查数据\n\r\n\r" + e1.Message);
            }
            finally
            {
                ex.Close();
                dictInventory = new Dictionary<string, int>();
                lstshijikucun.Items.Clear();
                lblxitongkucun.Text = string.Empty;
            }
        }
        /// <summary>
        /// 根据execl路径，添加实际库存到内存
        /// </summary>
        /// <param name="fileName"></param>
        private void AddShiJiKuCun(string fileName)
        {
            ExcelHelp ex = new ExcelHelp();
            try
            {
                ex.Open(fileName);
                ex.GetSheetByNumber(1);
                int hangshu = ex.ws.UsedRange.Rows.Count;
                for (int i = 1; i <= hangshu; i++)
                {
                    string kuanHao = Convert.ToString(ex.GetCellValue(i, 1)).Trim().ToUpper();
                    int kuanHaoCount = 0;
                    //B列是数字，A列有值
                    if ((Int32.TryParse(Convert.ToString(ex.GetCellValue(i, 2)), out kuanHaoCount)) == true && string.IsNullOrEmpty(kuanHao) == false)
                    {
                        if (dictInventory.ContainsKey(kuanHao))
                        {
                            int count = Convert.ToInt32(dictInventory[kuanHao]);
                            dictInventory.Remove(kuanHao);
                            dictInventory.Add(kuanHao, kuanHaoCount + count);
                        }
                        else
                        {
                            dictInventory.Add(kuanHao, kuanHaoCount);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                ex.Close();
            }
        }

        /// <summary>
        /// 根据execl路径，减去系统库存到内存
        /// </summary>
        /// <param name="fileName"></param>
        private void SubXiTongKuCun(string fileName)
        {
            ExcelHelp ex = new ExcelHelp();
            try
            {
                ex.Open(fileName);
                ex.GetSheetByNumber(1);
                int hangshu = ex.ws.UsedRange.Rows.Count;
                for (int i = 1; i <= hangshu; i++)
                {
                    string kuanHao = Convert.ToString(ex.GetCellValue(i, 2)).Trim().ToUpper();
                    int kuanHaoCount = 0;
                    //P列是数字，B列有值
                    if ((Int32.TryParse(Convert.ToString(ex.GetCellValue(i, 16)), out kuanHaoCount)) == true
                        && string.IsNullOrEmpty(kuanHao) == false
                        && kuanHao.Equals("") == false
                        && kuanHao.Equals("TOTAL") == false
                        )
                    {
                        if (dictInventory.ContainsKey(kuanHao))
                        {
                            int count = Convert.ToInt32(dictInventory[kuanHao]);
                            dictInventory.Remove(kuanHao);
                            dictInventory.Add(kuanHao, count - kuanHaoCount);
                        }
                        else
                        {
                            dictInventory.Add(kuanHao, -kuanHaoCount);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                ex.Close();
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            dictInventory = new Dictionary<string, int>();
            lstshijikucun.Items.Clear();
            lblxitongkucun.Text = string.Empty;
        }
    }
}
