using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Kombinasyon
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OpenFileDialog file;
        public DataTable getirExcelTablo()
        {
            OleDbConnection baglanti = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file.FileName + ";Extended Properties=Excel 12.0;");
            baglanti.Open();
            string query = "select * from [Sayfa1$]";
            OleDbDataAdapter oAdp = new OleDbDataAdapter(query, baglanti);
            DataTable dt = new DataTable();
            oAdp.Fill(dt);
            dataGridView1.DataSource = dt;
            return dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            file = new OpenFileDialog();
            file.Filter = "Excel Dosyası |*.xlsx";
            file.ShowDialog();

            DataTable dt = getirExcelTablo();

            List<List<object>> okunanList = new List<List<object>>();
            List<object> objectArr1 = new List<object>();
            List<object> objectArr2 = new List<object>();
            List<object> objectArr3 = new List<object>();
            List<object> objectArr4 = new List<object>();
            List<object> objectArr5 = new List<object>();

            for (int a = 0; a < dt.Rows.Count; a++)
            {

                if (dt.Rows[0]["T_GK_1"].ToString() != "")
                {
                    if (dt.Rows[a]["T_GK_1"].ToString() != "")
                    {

                        objectArr1.Add(dt.Rows[a]["T_GK_1"].ToString());
                    }
                }
                if (dt.Rows[0]["T_GK_2"].ToString() != "")
                {

                    if (dt.Rows[a]["T_GK_2"].ToString() != "")
                    {
                        objectArr2.Add(dt.Rows[a]["T_GK_2"].ToString());
                    }
                }
                if (dt.Rows[0]["T_GK_3"].ToString() != "")
                {
                    if (dt.Rows[a]["T_GK_3"].ToString() != "")
                    {
                        objectArr3.Add(dt.Rows[a]["T_GK_3"].ToString());
                    }
                }
                if (dt.Rows[0]["T_GK_4"].ToString() != "")
                {
                    if (dt.Rows[a]["T_GK_4"].ToString() != "")
                    {
                        objectArr4.Add(dt.Rows[a]["T_GK_4"].ToString());
                    }
                }

                if (dt.Rows[0]["T_GK_5"].ToString() != "")
                {
                    if (dt.Rows[a]["T_GK_5"].ToString() != "")
                    {
                        objectArr5.Add(dt.Rows[a]["T_GK_5"].ToString());
                    }
                }
            }

            if (objectArr1.Count != 0) okunanList.Add(objectArr1);
            if (objectArr2.Count != 0) okunanList.Add(objectArr2);
            if (objectArr3.Count != 0) okunanList.Add(objectArr3);
            if (objectArr4.Count != 0) okunanList.Add(objectArr4);
            if (objectArr5.Count != 0) okunanList.Add(objectArr5);

            List<object[]> kombList = new List<object[]>(); // Kombinasyon verilerinin tutulduğu liste

            kombList = Kombinasyon(okunanList, kombList, 0);
            DataTable dt2 = new DataTable();
            dt2 = ConvertListToDataTable(kombList);
            dataGridView2.DataSource = dt2;
        }

        private static List<object[]> Kombinasyon(List<List<object>> list, List<object[]> newList, params int[] loopInd)
        {
            if (loopInd.Length <= list.Count) // loopInd -> Döngü Sayısı. Verilen listenin boyutunu geçmemeli.
            {
                int currentCount = list[loopInd.Length - 1].Count;

                while (loopInd[loopInd.Length - 1] < currentCount) // i1<list[0] , i2<list[1] , i3<list[2]
                {
                    Kombinasyon(list, newList, loopInd.Concat(new[] { 0 }).ToArray()); // iç döngü için 
                    loopInd[loopInd.Length - 1]++; // i++, i2++ , i3++ ...
                }
            }
            else
            {
                int j = 0;
                object[] temp = loopInd.Take(loopInd.Length - 1).Select(i => list[j++][i]).ToArray();
                newList.Add(temp);
            }
            return newList;
        }

        public DataTable ConvertListToDataTable(List<object[]> list)
        {
            DataTable table = new DataTable();
            int columns = 0;
            foreach (var array in list)
            {
                if (array.Length > columns)
                {
                    columns = array.Length;
                }
            }
            // Column Ekle.
            for (int i = 0; i < columns; i++)
            {
                table.Columns.Add();
            }
            // Row Ekle.
            foreach (var array in list)
            {
                table.Rows.Add(array);
            }
            return table;
        }
    }
}
