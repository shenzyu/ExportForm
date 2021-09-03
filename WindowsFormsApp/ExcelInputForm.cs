using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace WindowsFormsApp
{
    public partial class ExcelInputForm : Form
    {
        private DataTable excelDataTable = new DataTable();
        private DataTable exportDataTable = new DataTable();

        public static int lengthLimit = 0;

        public ExcelInputForm()
        {
            InitializeComponent();
        }

        private void btn_Office_Click(object sender, EventArgs e)
        {


            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Files|*.xls;*.xlsx",              // 设定打开的文件类型
                                                            //openFileDialog.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;                       // 打开app对应的路径
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)  // 打开桌面
            };

            // 如果选定了文件
            string filePath = "";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {


                // 取得文件路径及文件名
                filePath = openFileDialog.FileName;
                excelDataTable.Clear();
                dataGridView1.DataSource = null;                       // 每次打开清空内容
                this.excelDataTable = ReadExcelToTable(filePath);      // 读出excel并放入datatable
                dataGridView1.DataSource = excelDataTable;        // 测试用, 输出到dataGridView
            }

        }
        private void button1_Click(object sender, EventArgs e)
        {

            UpDownBase up = (UpDownBase)numUpDown;
            if (!string.IsNullOrEmpty(up.Text))
            {
                string text = up.Text;
                int.TryParse(text, out lengthLimit);
            }
            if (lengthLimit < 1)
            {
                MessageBox.Show("限制长度需要大于0!");
                return;
            }

            List<InputModel> excelData = new List<InputModel>();
            try
            {
                for (int j = 1; j <= dataGridView1.RowCount - 2; j++)
                {
                    InputModel item = new InputModel();
                    for (int i = 0; i <= dataGridView1.ColumnCount - 1; i++)
                    {
                        string title = dataGridView1.Rows[0].Cells[i].Value.ToString();
                        switch (title)
                        {
                            case "购方企业名称":
                                item.FormName = dataGridView1.Rows[j].Cells[i].Value.ToString();
                                break;
                            case "开票日期":
                                item.OpenDate = dataGridView1.Rows[j].Cells[i].Value.ToString();
                                break;
                            case "商品名称":
                                item.Name = dataGridView1.Rows[j].Cells[i].Value.ToString();
                                break;
                            case "规格":
                                item.Spec = dataGridView1.Rows[j].Cells[i].Value.ToString();
                                break;
                            case "单位":
                                item.Unit = dataGridView1.Rows[j].Cells[i].Value.ToString();
                                break;
                            case "数量":
                                item.Num = dataGridView1.Rows[j].Cells[i].Value.ToString();
                                break;
                        }
                    }
                    excelData.Add(item);
                }

                if (excelDataTable.Rows.Count < 1)
                {
                    MessageBox.Show("表格数据为空,请确认是否已经导入数据!");
                    return;
                }
                exportDataTable.Clear();
                exportDataTable = GetDataTable(excelData);
                if (exportDataTable.Rows.Count > 0)
                {
                    dataGridView1.DataSource = null;
                    dataGridView1.DataSource = exportDataTable;
                }
                else
                {
                    MessageBox.Show("生成失败,请联系开发人员!");

                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public static DataTable GetDataTable(List<InputModel> excelData)
        {

            //表格绘制
            DataTable exportDataTable = new DataTable();
            DataColumn dc = exportDataTable.Columns.Add("Id", Type.GetType("System.String"));
            dc = exportDataTable.Columns.Add("Name", Type.GetType("System.String"));
            dc = exportDataTable.Columns.Add("Spec", Type.GetType("System.String"));
            dc = exportDataTable.Columns.Add("Unit", Type.GetType("System.String"));
            dc = exportDataTable.Columns.Add("Num", Type.GetType("System.String"));
            dc = exportDataTable.Columns.Add("Memo", Type.GetType("System.String"));

            //1.先根据出库单位分组
            var valueListGroup = excelData.GroupBy(x => new { x.FormName,x.OpenDate })
                      .Select(group => group).ToList();
            valueListGroup.ForEach(item =>
            {
                List<List<InputModel>> valueList = SplitList(item.ToList());
                valueList.ForEach(oList =>
                {
                    //oList即为一个表的数据
                    //先添加两个空行
                    DataTable itemDataTable = new DataTable();
                    DataColumn dcitem = itemDataTable.Columns.Add("Id", Type.GetType("System.String"));
                    dcitem = itemDataTable.Columns.Add("Name", Type.GetType("System.String"));
                    dcitem = itemDataTable.Columns.Add("Spec", Type.GetType("System.String"));
                    dcitem = itemDataTable.Columns.Add("Unit", Type.GetType("System.String"));
                    dcitem = itemDataTable.Columns.Add("Num", Type.GetType("System.String"));
                    dcitem = itemDataTable.Columns.Add("Memo", Type.GetType("System.String"));

                    itemDataTable.Rows.InsertAt(itemDataTable.NewRow(), 0);
                    itemDataTable.Rows.InsertAt(itemDataTable.NewRow(), 1);
                    //标题
                    DataRow titleDr = itemDataTable.NewRow();
                    titleDr["Id"] = "雪 海 梅 乡 食 品 出 库 单";
                    titleDr["Spec"] = null;
                    titleDr["Unit"] = null;
                    itemDataTable.Rows.Add(titleDr);
                    //空行
                    itemDataTable.Rows.InsertAt(itemDataTable.NewRow(), 3);
                    //信息行
                    DataRow messageDr = itemDataTable.NewRow();
                    messageDr["Id"] = "发货日期：";
                    messageDr["Name"] = oList.FirstOrDefault().OpenDate;
                    messageDr["Spec"] = "单位:";
                    messageDr["Unit"] = oList.FirstOrDefault().FormName;
                    
                    itemDataTable.Rows.Add(messageDr);


                    DataRow headDr = itemDataTable.NewRow();
                    headDr["Id"] = "序号";
                    headDr["Name"] = "品名";
                    headDr["Spec"] = "规格";
                    headDr["Unit"] = "单位";
                    headDr["Num"] = "数量";
                    headDr["Memo"] = "备注";
                   
                    itemDataTable.Rows.Add(headDr);
                    
                    //数据
                    var sum = 0;
                    for (int i = 0; i < lengthLimit; i++)
                    {
                        if (oList.Count > (i))
                        {
                            InputModel inputModel = oList[i];
                            DataRow itemDr = itemDataTable.NewRow();
                            itemDr["Id"] = (i + 1).ToString();
                            itemDr["Name"] = inputModel.Name;
                            itemDr["Spec"] = inputModel.Spec;
                            itemDr["Unit"] = inputModel.Unit;
                            itemDr["Num"] = inputModel.Num;
                            itemDr["Memo"] = string.Empty;
                            itemDataTable.Rows.Add(itemDr);
                            Int32.TryParse(inputModel.Num, out int num);
                            sum = sum + num;
                        }
                        else
                        {
                            DataRow itemDr = itemDataTable.NewRow();
                            itemDr["Id"] = (i + 1).ToString();
                            itemDataTable.Rows.Add(itemDr);
                        }
                    }
                    //添加合计行
                    DataRow sumDr = itemDataTable.NewRow();
                    sumDr["Name"] = "合   计:";
                    sumDr["Num"] = sum;
                    itemDataTable.Rows.Add(sumDr);

                    //空行
                    itemDataTable.Rows.Add(itemDataTable.NewRow());
                    //经办人
                    DataRow HandlerDr = itemDataTable.NewRow();
                    HandlerDr["Name"] = "经办人：";
                    itemDataTable.Rows.Add(HandlerDr);
                    //空行
                    itemDataTable.Rows.Add(itemDataTable.NewRow());
                    //注:
                    DataRow needDr = itemDataTable.NewRow();
                    needDr["Id"] = "注：第一联记帐联；第二联仓库联.";
                    itemDataTable.Rows.Add(needDr);
                    //空行
                    //itemDataTable.Rows.Add(itemDataTable.NewRow());
                    //添加到导出表
                    foreach (DataRow dr in itemDataTable.Rows)
                    {
                        
                        exportDataTable.ImportRow(dr);
                       
                    }
                });
            });

            return exportDataTable;
        }

        public static List<List<InputModel>> SplitList(List<InputModel> list)
        {
            var clsListInputModel = new List<List<InputModel>>();

            if (list.Count < lengthLimit)
            {
                clsListInputModel.Add(list);
            }
            else
            {
                var count = Math.Ceiling((decimal)list.Count / lengthLimit);
                for (int i = 0; i < count; i++)
                {
                    var itemList = new List<InputModel>();
                    for (int j = (i) * lengthLimit; j < list.Count; j++)
                    {
                        if (j <= ((i + 1) * lengthLimit))
                        {
                            itemList.Add(list[j]);
                        }
                    }
                    clsListInputModel.Add(itemList);
                }
            }
            return clsListInputModel;

        }
        public static XSSFWorkbook BuildWorkbook(DataTable dt)
        {
            var book = new XSSFWorkbook();
            ISheet sheet = book.CreateSheet("Sheet1");


            ICellStyle cellStyle = book.CreateCellStyle();

            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;

            //Data Rows
            HashSet<int> mergeRowIndex = new HashSet<int>();
            HashSet<int> mergeRowIndex1 = new HashSet<int>();
            HashSet<int> mergeRowIndex2 = new HashSet<int>();
            int borderInt = -1;
            Boolean isStart = false;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow drow = sheet.CreateRow(i);
                if (isStart)
                {
                    if (i - borderInt > 12)
                    {
                        borderInt = -1;
                        isStart = false;
                    }
                }


                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = drow.CreateCell(j, CellType.String);
                    if (dt.Rows[i][j].ToString() == "序号")
                    {
                        borderInt = i;
                        isStart = true;
                    }
                    if (dt.Rows[i][j].ToString() == "雪 海 梅 乡 食 品 出 库 单")
                    {
                        mergeRowIndex.Add(i);
                    }
                    if (dt.Rows[i][j].ToString() == "注：第一联记帐联；第二联仓库联.")
                    {
                        mergeRowIndex1.Add(i);
                    }
                    if (dt.Rows[i][j].ToString() == "发货日期：")
                    {
                        mergeRowIndex2.Add(i);
                    }
                    if (i >= borderInt &&   i- borderInt < 12 && isStart)
                    {
                        cell.CellStyle = cellStyle;
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                    {
                        cell.SetCellValue(dt.Rows[i][j].ToString());

                    }
                }

            }

            //自动列宽
            for (int i = 0; i <= dt.Columns.Count; i++)
                sheet.AutoSizeColumn(i, true);

            //合并单元格//设置style
            ICellStyle mergeStyle = book.CreateCellStyle();
            mergeStyle.VerticalAlignment = VerticalAlignment.Center;
            mergeStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;

            ICellStyle mergeCenterStyle = book.CreateCellStyle();
            mergeCenterStyle.VerticalAlignment = VerticalAlignment.Center;
            mergeCenterStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

            //合并操作
            foreach (int rowIndex in mergeRowIndex)
            {
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 2));
                IRow row = sheet.CreateRow(rowIndex);
                ICell cell = row.CreateCell(0);
                cell.CellStyle = mergeStyle;
                cell.SetCellValue("雪 海 梅 乡 食 品 出 库 单");
                
            }
            //合并操作
            foreach (int rowIndex in mergeRowIndex1)
            {
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 2));
                IRow row = sheet.CreateRow(rowIndex);
                ICell cell = row.CreateCell(0);
                cell.CellStyle = mergeStyle;
                cell.SetCellValue("注：第一联记帐联；第二联仓库联.");

            }
            //合并操作
            int idNo = 1;
            foreach (int rowIndex in mergeRowIndex2)
            {
                var value = sheet.GetRow(rowIndex).GetCell(3).ToString();
                var date = sheet.GetRow(rowIndex).GetCell(1).ToString();
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 3, 4));

                IRow row = sheet.CreateRow(rowIndex);
                ICell cell = row.CreateCell(3);
                ICell cellString = row.CreateCell(0);
                ICell cellDate = row.CreateCell(1);
                ICell formNameCell = row.CreateCell(2);
                ICell IdCell = row.CreateCell(5);

                cell.CellStyle = mergeStyle;
                cellString.CellStyle = mergeStyle;
                cellDate.CellStyle = mergeStyle;
                IdCell.CellStyle = mergeCenterStyle;
                formNameCell.CellStyle = mergeStyle;

                cell.SetCellValue(value);
                cellString.SetCellValue("发货日期：");
                cellDate.SetCellValue(date);
                formNameCell.SetCellValue("单位:");
                IdCell.SetCellValue(idNo);
                idNo++;
            }
            return book;
        }
        private static DataTable ReadExcelToTable(string path)
        {
            try
            {
                // 连接字符串(Office 07及以上版本 不能出现多余的空格 而且分号注意)
                string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";
                // 连接字符串(Office 07以下版本, 基本上上面的连接字符串就可以了) 
                //string connstring = Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";

                using (OleDbConnection conn = new OleDbConnection(connstring))
                {
                    conn.Open();
                    // 取得所有sheet的名字
                    DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });
                    // 取得第一个sheet的名字
                    string firstSheetName = sheetsName.Rows[0][2].ToString();

                    // 查询字符串 
                    string sql = string.Format("SELECT * FROM [{0}]", firstSheetName);

                    // OleDbDataAdapter是充当 DataSet 和数据源之间的桥梁，用于检索和保存数据
                    OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);

                    // DataSet是不依赖于数据库的独立数据集合
                    DataSet set = new DataSet();

                    // 使用 Fill 将数据从数据源加载到 DataSet 中
                    ada.Fill(set);

                    return set.Tables[0];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }

        }

        private void ExcelInputForm_Load(object sender, EventArgs e)
        {
            numUpDown.Value = 10;

        }

        private void btnDownload_Click(object sender, EventArgs e)
        {
             
            var newBook = BuildWorkbook(exportDataTable);
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Files|*.xlsx",              // 设定打开的文件类型
                                                            //openFileDialog.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;                       // 打开app对应的路径
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)  // 打开桌面
            };
            string path = "";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                path = saveFileDialog.FileName;
            }
            else
            {
                return;
            }
            using (var fs = File.OpenWrite(path))
            {

                newBook.Write(fs);
            }
        }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
        }
    }
}
