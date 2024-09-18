using LinqLabs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyHomeWork
{
    public partial class Frm作業_2 : Form
    {
        public Frm作業_2()
        {
            InitializeComponent();
            this.productPhotoTableAdapter1.Fill(this.adwDataSet1.ProductPhoto);

            var q1 = from p in this.adwDataSet1.ProductPhoto
                     select p.ModifiedDate.Year;
            this.comboBox3.DataSource = q1.Distinct().ToList();
        }

        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            var q = from o in this.adwDataSet1.ProductPhoto
                    select o;
            this.bindingSource1.DataSource = q.ToList();
            this.dataGridView1.DataSource = this.bindingSource1;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            var q = from o in this.adwDataSet1.ProductPhoto
                    where o.ModifiedDate > dateTimePicker1.Value && o.ModifiedDate < dateTimePicker2.Value
                    select o;
            this.bindingSource1.DataSource = q.ToList();
            this.dataGridView1.DataSource = this.bindingSource1;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            int year;
            int.TryParse(this.comboBox3.Text, out year);
            var q = from o in this.adwDataSet1.ProductPhoto
                    where o.ModifiedDate.Year == year
                    select o;

            this.bindingSource1.DataSource = q.ToList();
            this.dataGridView1.DataSource = this.bindingSource1;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            int selectedYear;
            string selectedSeason = comboBox2.Text;

            int.TryParse(this.comboBox3.Text, out selectedYear);

            var q = from o in this.adwDataSet1.ProductPhoto
                    where o.ModifiedDate.Year == selectedYear && Season(o.ModifiedDate) == selectedSeason

                    select o;



            this.bindingSource1.DataSource = q.ToList();
            this.dataGridView1.DataSource = this.bindingSource1;

            MessageBox.Show("共" + q.Count() + " 筆項目");
        }
        public static string Season(DateTime source)
        {
            switch (source.Month)
            {
                case 1:
                case 2:
                case 3:
                    return "第一季";
                case 4:
                case 5:
                case 6:
                    return "第二季";
                case 7:
                case 8:
                case 9:
                    return "第三季";
                case 10:
                case 11:
                case 12:
                    return "第四季";

                default:
                    return "";

            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            this.lblMaster.Text = "ProductPhoto";

            int selectedYear;
            int.TryParse(this.comboBox3.Text, out selectedYear);
            string selectedSeason = comboBox2.Text;

            var q = from o in this.adwDataSet1.ProductPhoto
                    where o.ModifiedDate.Year == selectedYear && Season(o.ModifiedDate) == selectedSeason
                    select o;

            this.bindingSource1.DataSource = q.ToList();
            this.dataGridView1.DataSource = this.bindingSource1;
            MessageBox.Show("共" + q.Count() + " 筆項目");
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            this.lblMaster.Text = "ProductPhoto";
            int year;
            int.TryParse(this.comboBox3.Text, out year);


            var q = from o in this.adwDataSet1.ProductPhoto
                    where o.ModifiedDate.Year == year
                    select o;

            this.bindingSource1.DataSource = q.ToList();
            this.dataGridView1.DataSource = this.bindingSource1;
        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {
            try
            {
                // 檢查當前選中的資料是否為空
                if (this.bindingSource1.Current == null)
                {
                    // 如果當前沒有選中的資料，清除 PictureBox 的圖片
                    this.pictureBox1.Image = null;
                    return;
                }

                // 取得當前選中的 ProductPhoto 資料列
                var currentRow = (ADWDataSet.ProductPhotoRow)this.bindingSource1.Current;

                //  LargePhoto 是二進制影像數據 (byte[])
                var largePhotoData = currentRow.LargePhoto;

                // 將二進制數據轉換成影像
                using (var ms = new System.IO.MemoryStream(largePhotoData))
                {
                    this.pictureBox1.Image = Image.FromStream(ms);
                }
            }
            catch
            {

            }
        }
    }
}
