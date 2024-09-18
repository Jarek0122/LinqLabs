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
    public partial class Frm作業_1 : Form
    {
        public Frm作業_1()
        {
            InitializeComponent();
            this.productsTableAdapter1.Fill(this.nwDataSet1.Products);
            this.ordersTableAdapter1.Fill(this.nwDataSet1.Orders);
            this.order_DetailsTableAdapter1.Fill(this.nwDataSet1.Order_Details);

            var q1 = from o in this.nwDataSet1.Orders
                     select o.OrderDate.Year;
            this.comboBox1.DataSource = q1.Distinct().ToList();

        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {
            try
            {
                this.lblDetails.Text = "Order details";

                var OrderID = ((NWDataSet.OrdersRow)this.bindingSource1.Current).OrderID;

                var q = from o in this.nwDataSet1.Order_Details
                        where o.OrderID == OrderID
                        select o;

                this.dataGridView2.DataSource = q.ToList();
            }
            catch
            {

            }
        }

        

        private void button14_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;


            System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(@"c:\windows");

            System.IO.FileInfo[] files = dir.GetFiles();

            var l = from log in dir.GetFiles()
                    where log.Extension.ToLower() == ".log"
                    select log;
            this.dataGridView1.DataSource = l.ToList();

        }

        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            this.lblMaster.Text = "Orders";
            this.dataGridView1.DataSource = this.nwDataSet1.Orders.ToList();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(@"c:\windows");

            System.IO.FileInfo[] files = dir.GetFiles();

            var l = from log in dir.GetFiles()
                    where log.CreationTime.Year == 2022
                    select log;
            this.dataGridView1.DataSource = l.ToList();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(@"c:\windows");

            System.IO.FileInfo[] files = dir.GetFiles();

            var l = from log in dir.GetFiles()
                    where log.Length > 1000000
                    select log;
            this.dataGridView1.DataSource = l.ToList();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(@"c:\windows");
            System.IO.FileInfo[] files = dir.GetFiles();
            this.dataGridView1.DataSource = files;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            int year;
            int.TryParse(this.comboBox1.Text, out year);

            var q2 = from o in this.nwDataSet1.Orders
                     where o.OrderDate.Year == year
                     select o;
            this.bindingSource1.DataSource = q2.ToList();
            this.dataGridView1.DataSource = this.bindingSource1;
        }

        int page = -1;
        int countPerPage = 10;

        private void button12_Click(object sender, EventArgs e)
        {
            int.TryParse(this.textBox1.Text, out countPerPage);

            page -= 1;

            this.dataGridView1.DataSource = this.nwDataSet1.Products.Skip(countPerPage * page).Take(countPerPage).ToList();
        }
        private void button13_Click(object sender, EventArgs e)
        {
            int.TryParse(this.textBox1.Text, out countPerPage);

            page += 1;
            this.dataGridView1.DataSource = this.nwDataSet1.Products.Skip(countPerPage * page).Take(countPerPage).ToList();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.lblMaster.Text = "Orders";

            int year = 1997;
            int.TryParse(this.comboBox1.Text, out year);

            var q = from o in this.nwDataSet1.Orders
                    where o.OrderDate.Year == year
                    select o;

            this.bindingSource1.DataSource = q.ToList();
            this.dataGridView1.DataSource = this.bindingSource1;
        }
    }
}
