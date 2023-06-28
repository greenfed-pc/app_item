using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Drawing.Printing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.CompilerServices;
using System.Data.OleDb;
using System.IO;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace app_item
{
    public partial class Form1 : Form
    {
        
        private SqlConnection sqlConnection = null;
        private SqlCommandBuilder sqlBuilder = null;
        private SqlDataAdapter sqlDataAdapter = null;
        private DataSet dataSet = null;

        private bool newRowAdding = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void LoadData ()
        {
            try
            {
                sqlDataAdapter = new SqlDataAdapter("select *,'Delete' as [Delete] from item ", sqlConnection);

                sqlBuilder = new SqlCommandBuilder(sqlDataAdapter);

                sqlBuilder.GetDeleteCommand();
                sqlBuilder.GetInsertCommand();
                sqlBuilder.GetUpdateCommand();

                dataSet = new DataSet();

                sqlDataAdapter.Fill(dataSet, "item");
                dataGridView1.DataSource = dataSet.Tables["item"];

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView1[14,i] = linkCell;
                }

            }
            catch(Exception ex) { MessageBox.Show(ex.Message, "Ошибка подключения!", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
        private void ReloadData()
        {
            try
            {
                dataSet.Tables.Clear();

                sqlDataAdapter.Fill(dataSet, "item");
                dataGridView1.DataSource = dataSet.Tables["item"];

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dataGridView1[14, i] = linkCell;
                }

            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Ошибка подключения!", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void eToolStripMenuItem_Click(object sender, EventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            //Книга.
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            //Вызываем нашу созданную эксельку.
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

            private void печатьToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Green\source\repos\app_item\item_db.mdf;Integrated Security=True");

            sqlConnection.Open();
            LoadData();

        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ReloadData();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex ==14)
                {
                    string task = dataGridView1.Rows[e.RowIndex].Cells[14].Value.ToString();
                    if (task == "Delete")
                    {
                        if (MessageBox.Show("Удалить строку таблицы?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            int rowIndex = e.RowIndex;
                            dataGridView1.Rows.RemoveAt(rowIndex);
                            dataSet.Tables["item"].Rows[rowIndex].Delete();

                            sqlDataAdapter.Update(dataSet, "item");
                        }
                       
                    }
                    else if (task =="Insert")
                    {
                        int rowIndex = dataGridView1.Rows.Count - 2;
                        DataRow row = dataSet.Tables["item"].NewRow();

                        // row[""]= dataGridView1.Rows[rowIndex].Cells[""].Value;
                        row["[Номер]"] = dataGridView1.Rows[rowIndex].Cells["[Номер]"].Value;
                        row["[Тип оргтехники]"] = dataGridView1.Rows[rowIndex].Cells["[Тип оргтехники]"].Value;
                        row["[Модель]"] = dataGridView1.Rows[rowIndex].Cells["[Модель]"].Value;
                        row["[Дата поступления]"] = dataGridView1.Rows[rowIndex].Cells["[Модель]"].Value;
                        row["[Гарантия до]"] = dataGridView1.Rows[rowIndex].Cells["[Гарантия до]"].Value;
                        row["[Стоимость]"] = dataGridView1.Rows[rowIndex].Cells["[Стоимость]"].Value;
                        row["[оставщик]"] = dataGridView1.Rows[rowIndex].Cells["[Поставщик]"].Value;
                        row["[ФИО_мат_ответ]"] = dataGridView1.Rows[rowIndex].Cells["[ФИО_мат_ответ]"].Value;
                        row["[Цех]"] = dataGridView1.Rows[rowIndex].Cells["[Цех]"].Value;
                        row["[Отдел]"] = dataGridView1.Rows[rowIndex].Cells["[Отдел]"].Value;
                        row["[Кабинет]"] = dataGridView1.Rows[rowIndex].Cells["[Кабинет]"].Value;
                        row["[ФИО_работника]"] = dataGridView1.Rows[rowIndex].Cells["[ФИО_работника]"].Value;
                        row["[На списание(-/+)]"] = dataGridView1.Rows[rowIndex].Cells["[На списание(-/+)]"].Value;

                        dataSet.Tables["item"].Rows.Add(row);
                        dataSet.Tables["item"].Rows.RemoveAt(dataSet.Tables["item"].Rows.Count - 1);
                        dataGridView1.Rows.RemoveAt(dataGridView1.Rows.Count - 2);
                        dataGridView1.Rows[e.RowIndex].Cells[14].Value = "Delete";

                        sqlDataAdapter.Update(dataSet, "item");
                        newRowAdding = false;

                    }
                    else if (task =="Update")
                    {
                        int r = e.RowIndex;
                        //dataSet.Tables["item"].Rows[r][""]= dataGridView1.Rows[r].Cells[""].Value;
                        dataSet.Tables["item"].Rows[r]["[Номер]"] = dataGridView1.Rows[r].Cells["[Номер]"].Value;
                        dataSet.Tables["item"].Rows[r]["[Тип оргтехники]"] = dataGridView1.Rows[r].Cells["[Тип оргтехники]"].Value;
                        dataSet.Tables["item"].Rows[r]["[Модель]"] = dataGridView1.Rows[r].Cells["[Модель]"].Value;
                        dataSet.Tables["item"].Rows[r]["[Дата поступления]"] = dataGridView1.Rows[r].Cells["[Дата поступления]"].Value;
                        dataSet.Tables["item"].Rows[r]["[Гарантия до]"] = dataGridView1.Rows[r].Cells["[Гарантия до]"].Value;
                        dataSet.Tables["item"].Rows[r]["[Стоимость]"] = dataGridView1.Rows[r].Cells["[Стоимость]"].Value;
                        dataSet.Tables["item"].Rows[r]["[Поставщик]"] = dataGridView1.Rows[r].Cells["[Поставщик]"].Value;
                        dataSet.Tables["item"].Rows[r]["[ФИО_мат_ответ]"] = dataGridView1.Rows[r].Cells["[ФИО_мат_ответ]"].Value;
                        dataSet.Tables["item"].Rows[r]["[Цех]"] = dataGridView1.Rows[r].Cells["[Цех]"].Value;
                        dataSet.Tables["item"].Rows[r]["[Отдел]"] = dataGridView1.Rows[r].Cells["[Отдел]"].Value;
                        dataSet.Tables["item"].Rows[r]["[Кабинет]"] = dataGridView1.Rows[r].Cells["[Кабинет]"].Value;
                        dataSet.Tables["item"].Rows[r]["[ФИО_работника]"] = dataGridView1.Rows[r].Cells["[ФИО_работника]"].Value;
                        dataSet.Tables["item"].Rows[r]["[На списание(-/+)]"] = dataGridView1.Rows[r].Cells["[На списание(-/+)]"].Value;

                        sqlDataAdapter.Update(dataSet, "item");
                        dataGridView1.Rows[e.RowIndex].Cells[14].Value = "Delete";
                    }
                }
               

            }
            catch { }


        }

        private void dataGridView1_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                if (newRowAdding == false)
                {
                    newRowAdding = true;
                    int lastRow = dataGridView1.RowCount - 2;

                    DataGridViewRow row = dataGridView1.Rows[lastRow];
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                    dataGridView1[14, lastRow] = linkCell;

                    row.Cells["Delete"].Value = "Insert";
                }

            }
            catch(Exception ex) { MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }


        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (newRowAdding == false)
                {
                    int rowIndex = dataGridView1.SelectedCells[0].RowIndex;

                    DataGridViewRow editinRow = dataGridView1.Rows[rowIndex];
                    DataGridViewLinkCell cell = new DataGridViewLinkCell();

                    dataGridView1[14, rowIndex] = cell;
                    editinRow.Cells["Delete"].Value = "Update";
                }

            }
            catch(Exception ex) { MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }

        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(column_keypress);
            if (dataGridView1.CurrentCell.ColumnIndex ==13)
            {
                TextBox textBox = e.Control as TextBox;
                if (textBox != null) { textBox.KeyPress += new KeyPressEventHandler(column_keypress); }
            }

        }
        private void column_keypress (object sender, KeyPressEventArgs e)
        {
            if ((!char.IsControl(e.KeyChar)) && char.IsDigit(e.KeyChar)) { e.Handled = false; }

           
        }
        //print docs
        private void printDocument_PrintPage(object sender, PrintPageEventArgs e) // Метод печати для printDocument
        {
            Bitmap bmp = new Bitmap(dataGridView1.Size.Width + 10, dataGridView1.Size.Height + 10);
            dataGridView1.DrawToBitmap(bmp, dataGridView1.Bounds);
            e.Graphics.DrawImage(bmp, 0, 0);
        }

        private void pDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PdfPTable pdfTable = new PdfPTable(dataGridView1.ColumnCount);
            pdfTable.DefaultCell.Padding = 5;
            pdfTable.WidthPercentage = 100;
            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
            pdfTable.DefaultCell.BorderWidth = 1;

            //Adding Header row
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                pdfTable.AddCell(cell);
            }

            //Adding DataRow
            //foreach (DataGridViewRow row in dataGridView1.Rows)
            //{
            //    foreach (DataGridViewCell cell in row.Cells)
            //    {
            //        pdfTable.AddCell(cell.Value.ToString());
            //    }
            //}
            int row = dataGridView1.Rows.Count;
            int cell2 = dataGridView1.Rows[1].Cells.Count;
            for (int i = 0; i < row - 1; i++)
            {
                for (int j = 0; j < cell2; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value == null)
                    {
                        //return directly
                        //return;
                        //or set a value for the empty data
                        dataGridView1.Rows[i].Cells[j].Value = "null";
                    }
                    pdfTable.AddCell(dataGridView1.Rows[i].Cells[j].Value.ToString());
                }
            }

            //Exporting to PDF
            string folderPath = @"D:\Log\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (FileStream stream = new FileStream(folderPath + "DataGridViewExport.pdf", FileMode.Create))
            {
                Document pdfDoc = new Document(PageSize.A2, 10f, 10f, 10f, 0f);
                PdfWriter.GetInstance(pdfDoc, stream);
                pdfDoc.Open();
                pdfDoc.Add(pdfTable);
                pdfDoc.Close();
                stream.Close();
            }
            MessageBox.Show("Done");



        }

        private void docToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void пользовательToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var application = new Microsoft.Office.Interop.Word.Application();
            var doc = new Microsoft.Office.Interop.Word.Document();

            doc = application.Documents.Add(Template: @"C:\Users\Green\Desktop\manual.docx");
            application.Visible = true;
            doc.SaveAs2(FileName: @"C:\Users\Green\Desktop\manual.docx");

        }
    }

}
