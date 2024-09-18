using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LinqLabs.作業
{
    public partial class Frm作業_3 : Form
    {
        private Random _random = new Random();
        private List<Student> students_scores;
        private List<Student> randomStudentScores;
        public Frm作業_3()
        {
            InitializeComponent();

            //hint
            students_scores = new List<Student>()
                                         {
                                            new Student{ Name = "aaa", Class = "CS_101", Chi = 80, Eng = 80, Math = 50, Gender = "Male" },
                                            new Student{ Name = "bbb", Class = "CS_102", Chi = 80, Eng = 80, Math = 100, Gender = "Male" },
                                            new Student{ Name = "ccc", Class = "CS_101", Chi = 60, Eng = 50, Math = 75, Gender = "Female" },
                                            new Student{ Name = "ddd", Class = "CS_102", Chi = 80, Eng = 70, Math = 85, Gender = "Female" },
                                            new Student{ Name = "eee", Class = "CS_101", Chi = 80, Eng = 80, Math = 50, Gender = "Female" },
                                            new Student{ Name = "fff", Class = "CS_102", Chi = 80, Eng = 80, Math = 80, Gender = "Female" },

                                          };
        }

        public class Student
        {
            public string Name { get; set; }
            public string Class { get; set; }
            public int Chi { get; set; }
            public int Eng { get; set; }
            public int Math { get; set; }
            public string Gender { get; set; }


        }

        private void button36_Click(object sender, EventArgs e)
        {
            //共幾個 學員成績 ?
            int total = students_scores.Count();

            MessageBox.Show("總共有 " + total + " 個學員成績");

            //--------------------------------------------------------------------
            // 找出 前面三個 的學員所有科目成績
            var Three = students_scores.Take(3);

            StringBuilder sbr = new StringBuilder();

            foreach (var student in Three)
            {
                sbr.AppendLine($"學員名稱: {student.Name}, 班級: {student.Class}, 國文: {student.Chi}, 英文: {student.Eng}, 數學: {student.Math}");
            }

            MessageBox.Show(sbr.ToString(), "前三名學員成績");

            //--------------------------------------------------------------------
            // 找出 後面兩個 的學員所有科目成績	
            var lastTwoStudents = students_scores.Skip(students_scores.Count - 2).Take(2);

            StringBuilder sbr1 = new StringBuilder();

            foreach (var student in lastTwoStudents)
            {
                sbr1.AppendLine($"學員名稱: {student.Name}, 班級: {student.Class}, 國文: {student.Chi}, 英文: {student.Eng}, 數學: {student.Math}");
            }

            MessageBox.Show(sbr1.ToString(), "最後兩名學員成績");

            //--------------------------------------------------------------------
            // 找出 Name 'aaa','bbb','ccc' 的學員國文英文科目成績
            var Students = students_scores
                                   .Where(s => new[] { "aaa", "bbb", "ccc" }.Contains(s.Name))
                                   .Select(s => new
                                   {
                                       s.Name,
                                       s.Chi,
                                       s.Eng
                                   });

            StringBuilder sbr2 = new StringBuilder();
            foreach (var student in Students)
            {
                sbr2.AppendLine($"學員名稱: {student.Name}, 國文: {student.Chi}, 英文: {student.Eng}");
            }
            MessageBox.Show(sbr2.ToString(), "指定學員國文英文科目成績");

            //--------------------------------------------------------------------
            // 找出學員 'bbb' 的成績
            var b = students_scores
                                  .Where(s => new[] { "bbb" }.Contains(s.Name))
                                  .Select(s => new
                                  {
                                      s.Name,
                                      s.Chi,
                                      s.Eng,
                                      s.Math
                                  });

            StringBuilder sbr3 = new StringBuilder();
            foreach (var student in b)
            {
                sbr3.AppendLine($"學員名稱: {student.Name}, 國文: {student.Chi}, 英文: {student.Eng}, 數學{student.Math}");
            }
            MessageBox.Show(sbr3.ToString(), "b學員成績");

            //--------------------------------------------------------------------
            // 找出除了 'bbb' 學員的學員的所有成績 ('bbb' 退學)
            var exceptB = students_scores
                            .Where(s => s.Name != "bbb");

            StringBuilder sbr4 = new StringBuilder();
            foreach (var student in exceptB)
            {
                sbr4.AppendLine($"學員名稱: {student.Name}, 班級: {student.Class}, 國文: {student.Chi}, 英文: {student.Eng}, 數學: {student.Math}");
            }
            MessageBox.Show(sbr4.ToString(), "除了 'bbb' 的學員成績");

            //--------------------------------------------------------------------
            // 找出 'aaa', 'bbb' 'ccc' 學員 國文數學兩科 科目成績  |				
            var Students1 = students_scores
                                   .Where(s => new[] { "aaa", "bbb", "ccc" }.Contains(s.Name))
                                   .Select(s => new
                                   {
                                       s.Name,
                                       s.Chi,
                                       s.Math
                                   });

            StringBuilder sbr5 = new StringBuilder();
            foreach (var student in Students1)
            {
                sbr5.AppendLine($"學員名稱: {student.Name}, 國文: {student.Chi}, 數學: {student.Math}");
            }
            MessageBox.Show(sbr5.ToString(), "指定學員國文數學科目成績");

            //--------------------------------------------------------------------
            // 數學不及格 ... 是誰
            var tol = students_scores
                      .Where(s => s.Math < 60);

            StringBuilder sbr6 = new StringBuilder();
            foreach (var student in tol)
            {
                sbr6.AppendLine($"學員名稱: {student.Name}, 班級: {student.Class}, 國文: {student.Chi}, 英文: {student.Eng}, 數學: {student.Math}");
            }
            MessageBox.Show(sbr6.ToString(), "數學不及格 ... 是誰");
            #region 搜尋 班級學生成績

            // 
            // 共幾個 學員成績 ?						

            // 找出 前面三個 的學員所有科目成績					
            // 找出 後面兩個 的學員所有科目成績					

            // 找出 Name 'aaa','bbb','ccc' 的學員國文英文科目成績						

            // 找出學員 'bbb' 的成績	                          

            // 找出除了 'bbb' 學員的學員的所有成績 ('bbb' 退學)	

            // 找出 'aaa', 'bbb' 'ccc' 學員 國文數學兩科 科目成績  |				
            // 數學不及格 ... 是誰 
            #endregion
        }

        private void button37_Click(object sender, EventArgs e)
        {
            //個人 sum, min, max, avg
            var studentStatistics = students_scores
                                   .Select(s => new
                                   {
                                       s.Name,
                                       TotalScore = s.Chi + s.Eng + s.Math,
                                       MinScore = new[] { s.Chi, s.Eng, s.Math }.Min(),
                                       MaxScore = new[] { s.Chi, s.Eng, s.Math }.Max(),
                                       AvgScore = new[] { s.Chi, s.Eng, s.Math }.Average()
                                   }).OrderByDescending(stat => stat.TotalScore);

            StringBuilder sbr = new StringBuilder();
            foreach (var stat in studentStatistics)
            {
                sbr.AppendLine($"學員名稱: {stat.Name}");
                sbr.AppendLine($"  總分: {stat.TotalScore}");
                sbr.AppendLine($"  最低分: {stat.MinScore}");
                sbr.AppendLine($"  最高分: {stat.MaxScore}");
                sbr.AppendLine($"  平均分: {stat.AvgScore:F2}"); // 格式化為兩位小數
                sbr.AppendLine();
            }
            MessageBox.Show(sbr.ToString(), "每個學生個人成績統計");

            //--------------------------------------------------------------------
            // 統計 每個學生個人成績 並排序
            StringBuilder sbr1 = new StringBuilder();
            foreach (var stat in studentStatistics)
            {
                sbr1.AppendLine($"學員名稱: {stat.Name}");
                sbr1.AppendLine($"  總分: {stat.TotalScore}");

            }
            MessageBox.Show(sbr1.ToString(), "統計每個學生個人成績且排序");


        }

        private void button33_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            // 生成 100 個隨機學生資料
            randomStudentScores = GRStudents(100);

            // 對學生進行統計
            var studentStatistics = randomStudentScores
                .Select(s => new
                {
                    s.Name,
                    TotalScore = s.Chi + s.Eng + s.Math,
                    MinScore = new[] { s.Chi, s.Eng, s.Math }.Min(),
                    MaxScore = new[] { s.Chi, s.Eng, s.Math }.Max(),
                    AvgScore = new[] { s.Chi, s.Eng, s.Math }.Average()
                });

            // 創建 DataTable
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("學員名稱", typeof(string));
            dataTable.Columns.Add("總分", typeof(int));
            dataTable.Columns.Add("最低分", typeof(int));
            dataTable.Columns.Add("最高分", typeof(int));
            dataTable.Columns.Add("平均分", typeof(double));

            // 將統計結果添加到 DataTable 中
            foreach (var stat in studentStatistics)
            {
                dataTable.Rows.Add(
                    stat.Name,
                    stat.TotalScore,
                    stat.MinScore,
                    stat.MaxScore,
                    stat.AvgScore
                );
            }
            // 顯示結果到 DataGridView
            dataGridView1.DataSource = dataTable;
        }

        private List<Student> GRStudents(int count)
        {
            var students = new List<Student>();
            var classes = new[] { "CS_101", "CS_102" };
            var genders = new[] { "Male", "Female" };

            for (int i = 1; i <= count; i++)
            {
                students.Add(new Student
                {
                    Name = $"Student_{i}",
                    Class = classes[_random.Next(classes.Length)],
                    Chi = _random.Next(0, 101), // 0 到 100 分
                    Eng = _random.Next(0, 101), // 0 到 100 分
                    Math = _random.Next(0, 101), // 0 到 100 分
                    Gender = genders[_random.Next(genders.Length)]
                });
            }
            return students;
        }
        private void button10_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            var productsByYearMonth = from order in dbContext.Orders
                                      join detail in dbContext.Order_Details on order.OrderID equals detail.OrderID
                                      join product in dbContext.Products on detail.ProductID equals product.ProductID
                                      where order.OrderDate.HasValue
                                      group product by new
                                      {
                                          Year = order.OrderDate.Value.Year,
                                          Month = order.OrderDate.Value.Month
                                      } into yearMonthGroup
                                      select new
                                      {
                                          Year = yearMonthGroup.Key.Year,
                                          Month = yearMonthGroup.Key.Month,
                                          Products = yearMonthGroup.ToList()
                                      };
            // ---- 顯示到 DataGridView ----
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("年份", typeof(int));
            dataTable.Columns.Add("月份", typeof(int));
            dataTable.Columns.Add("產品名稱", typeof(string));
            dataTable.Columns.Add("單價", typeof(decimal));

            foreach (var yearMonthGroup in productsByYearMonth)
            {
                foreach (var product in yearMonthGroup.Products)
                {
                    dataTable.Rows.Add(
                        yearMonthGroup.Year,
                        yearMonthGroup.Month,
                        product.ProductName,
                        product.UnitPrice
                    );
                }
            }

            // 綁定 DataTable 到 DataGridView
            dataGridView1.DataSource = dataTable;
            // ---- 顯示到 TreeView ----
            treeView1.Nodes.Clear(); // 清空之前的節點

            foreach (var yearMonthGroup in productsByYearMonth)
            {
                // 創建年份月份根節點
                TreeNode yearMonthNode = new TreeNode($"年份: {yearMonthGroup.Year}, 月份: {yearMonthGroup.Month}");

                foreach (var product in yearMonthGroup.Products)
                {
                    // 每個產品作為子節點
                    TreeNode productNode = new TreeNode($"{product.ProductName} - {product.UnitPrice:C2}");
                    yearMonthNode.Nodes.Add(productNode);
                }

                // 將年份月份節點添加到 TreeView
                treeView1.Nodes.Add(yearMonthNode);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // split => 分成 三群 '待加強'(60~69) '佳'(70~89) '優良'(90~100)
            // print 每一群是哪幾個 ? (每一群 sort by 分數 descending)

            dataGridView1.DataSource = null;
            var studentStatistics = randomStudentScores
                                     .Select(s => new
                                     {
                                         s.Name,
                                         TotalScore = s.Chi + s.Eng + s.Math,
                                         MinScore = new[] { s.Chi, s.Eng, s.Math }.Min(),
                                         MaxScore = new[] { s.Chi, s.Eng, s.Math }.Max(),
                                         AvgScore = new[] { s.Chi, s.Eng, s.Math }.Average()
                                     });

            var groupedStudents = studentStatistics
                                   .GroupBy(stat =>
                                   {

                                       if (stat.TotalScore >= 90 && stat.TotalScore <= 100) return "優良";
                                       else if (stat.TotalScore >= 70 && stat.TotalScore <= 89) return "佳";
                                       else if (stat.TotalScore >= 60 && stat.TotalScore <= 69) return "待加強";
                                       return "不及格";
                                   })
                                   .Select(group => new
                                   {
                                       Category = group.Key,
                                       Students = group.OrderByDescending(stat => stat.TotalScore).ToList()
                                   });

            // ---- DataGridView1 顯示 ----

            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("群組", typeof(string));
            dataTable.Columns.Add("學員名稱", typeof(string));
            dataTable.Columns.Add("總分", typeof(int));

            foreach (var group in groupedStudents)
            {
                foreach (var student in group.Students)
                {
                    dataTable.Rows.Add(group.Category, student.Name, student.TotalScore);
                }
            }
            dataGridView1.DataSource = dataTable;

            // ---- TreeView1 顯示 ----
            treeView1.Nodes.Clear(); // 清空之前的節點

            foreach (var group in groupedStudents)
            {
                // 創建群組節點
                TreeNode groupNode = new TreeNode(group.Category);

                foreach (var student in group.Students)
                {
                    // 每個學生作為該群組的子節點
                    TreeNode studentNode = new TreeNode($"{student.Name} - 總分: {student.TotalScore}");
                    groupNode.Nodes.Add(studentNode);
                }

                // 將群組節點添加到 TreeView
                treeView1.Nodes.Add(groupNode);
            }

            // 展開 TreeView 的所有節點
            treeView1.ExpandAll();
        }

        private void button38_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(@"c:\windows");

            var groupedFiles = dir.GetFiles()
                .GroupBy(file =>
                {
                    if (file.Length < 1024 * 100) // 小於100KB
                        return "小檔案 (小於100KB)";
                    else if (file.Length >= 1024 * 100 && file.Length <= 1024 * 1024) // 100KB到1MB
                        return "中等檔案 (100KB到1MB)";
                    else
                        return "大檔案 (大於1MB)";
                })
                .Select(group => new
                {
                    Category = group.Key,
                    Files = group
                });

            // 創建 DataTable 來顯示檔案群組及其詳細信息
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("檔案群組", typeof(string));
            dataTable.Columns.Add("檔案名稱", typeof(string));
            dataTable.Columns.Add("檔案大小 (KB)", typeof(double));
            dataTable.Columns.Add("建立日期", typeof(DateTime));

            // 將群組和檔案信息填充到 DataTable
            foreach (var group in groupedFiles)
            {
                foreach (var file in group.Files)
                {
                    dataTable.Rows.Add(
                        group.Category,
                        file.Name,
                        Math.Round(file.Length / 1024.0, 2), // KB 單位
                        file.CreationTime
                    );
                }
            }
            // 將 DataTable 綁定到 DataGridView1
            dataGridView1.DataSource = dataTable;

            // ---- TreeView1 顯示 ----
            treeView1.Nodes.Clear(); // 清空之前的節點
            foreach (var group in groupedFiles)
            {
                // 創建群組節點
                TreeNode groupNode = new TreeNode(group.Category);

                foreach (var file in group.Files)
                {
                    // 每個檔案作為該群組的子節點
                    TreeNode fileNode = new TreeNode($"{file.Name} - {Math.Round(file.Length / 1024.0, 2)} KB");
                    groupNode.Nodes.Add(fileNode);
                }

                // 將群組節點添加到 TreeView
                treeView1.Nodes.Add(groupNode);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(@"c:\windows");

            var groupedFiles = dir.GetFiles()
                .GroupBy(file => file.CreationTime.Year)
                .Select(group => new
                {
                    Year = group.Key,
                    Files = group
                });

            // 創建 DataTable 來顯示檔案群組及其詳細信息
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("建立年份", typeof(int));
            dataTable.Columns.Add("檔案名稱", typeof(string));
            dataTable.Columns.Add("檔案大小 (KB)", typeof(double));
            dataTable.Columns.Add("建立日期", typeof(DateTime));

            // 將群組和檔案信息填充到 DataTable
            foreach (var group in groupedFiles)
            {
                foreach (var file in group.Files)
                {
                    dataTable.Rows.Add(
                        group.Year,
                        file.Name,
                        Math.Round(file.Length / 1024.0, 2), // KB 單位
                        file.CreationTime
                    );
                }
            }

            // 將 DataTable 綁定到 DataGridView1
            dataGridView1.DataSource = dataTable;

            // ---- TreeView1 顯示 ----
            treeView1.Nodes.Clear(); // 清空之前的節點
            foreach (var group in groupedFiles)
            {
                // 創建群組節點
                TreeNode groupNode = new TreeNode(group.Year.ToString());

                foreach (var file in group.Files)
                {
                    // 每個檔案作為該群組的子節點
                    TreeNode fileNode = new TreeNode($"{file.Name} - {Math.Round(file.Length / 1024.0, 2)} KB");
                    groupNode.Nodes.Add(fileNode);
                }

                // 將群組節點添加到 TreeView
                treeView1.Nodes.Add(groupNode);
            }
        }

        NorthwindEntities dbContext = new NorthwindEntities();

        private void button8_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;

            // 計算每個類別的平均價格
            var query = from p in dbContext.Products
                        group p by p.Category.CategoryName into g
                        select new
                        {
                            CategoryName = g.Key,
                            AvgUnitPrice = g.Average(p => p.UnitPrice)
                        };

            // 將查詢結果轉換為明確的類型
            var categorizedPrices = query.ToList().Select(item => new
            {
                item.CategoryName,
                PriceCategory = GetPriceCategory((decimal)(item.AvgUnitPrice ?? 0m)),
                AvgUnitPrice = item.AvgUnitPrice, // 確保 AvgUnitPrice 的類型是 decimal

            }).ToList();

            // 將結果顯示到 DataGridView
            dataGridView1.DataSource = categorizedPrices;

            // ---- TreeView1 顯示 ----
            treeView1.Nodes.Clear(); // 清空之前的節點

            // 根據分類顯示在 TreeView 中
            foreach (var category in categorizedPrices)
            {
                // 創建根節點
                TreeNode categoryNode = new TreeNode(category.CategoryName);

                // 創建子節點，顯示價格範疇
                TreeNode priceNode = new TreeNode($"{category.PriceCategory} - 平均價格: {category.AvgUnitPrice:C2}");
                categoryNode.Nodes.Add(priceNode);

                // 將根節點添加到 TreeView
                treeView1.Nodes.Add(categoryNode);
            }
        }
        private string GetPriceCategory(decimal avgUnitPrice)
        {

            if (avgUnitPrice < 30) // 假設低於 $30 為低價
                return "低價";
            else if (avgUnitPrice >= 30 && avgUnitPrice <= 50) // $30  到 $50 為中價
                return "中價";
            else // 高於 $50 為高價
                return "高價";
        }

        private void button15_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;

            // 根據年份分組產品
            var productsByYear = from order in dbContext.Orders
                                 join detail in dbContext.Order_Details on order.OrderID equals detail.OrderID
                                 join product in dbContext.Products on detail.ProductID equals product.ProductID
                                 group product by order.OrderDate.Value.Year into yearGroup
                                 select new
                                 {
                                     Year = yearGroup.Key,
                                     Products = yearGroup.ToList()
                                 };

            // ---- 顯示到 DataGridView ----
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("年份", typeof(int));
            dataTable.Columns.Add("產品名稱", typeof(string));
            dataTable.Columns.Add("單價", typeof(decimal));

            foreach (var yearGroup in productsByYear)
            {
                foreach (var product in yearGroup.Products)
                {
                    dataTable.Rows.Add(
                        yearGroup.Year,
                        product.ProductName,
                        product.UnitPrice
                    );
                }
            }

            // 綁定 DataTable 到 DataGridView
            dataGridView1.DataSource = dataTable;

            // ---- 顯示到 TreeView ----
            treeView1.Nodes.Clear(); // 清空之前的節點

            foreach (var yearGroup in productsByYear)
            {
                // 創建年份根節點
                TreeNode yearNode = new TreeNode($"年份: {yearGroup.Year}");

                foreach (var product in yearGroup.Products)
                {
                    // 每個產品作為子節點
                    TreeNode productNode = new TreeNode($"{product.ProductName} - {product.UnitPrice:C2}");
                    yearNode.Nodes.Add(productNode);
                }

                // 將年份節點添加到 TreeView
                treeView1.Nodes.Add(yearNode);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            // 根據年份和月份分組產品，計算每組的總銷售金額
            var productsByYearMonth = from order in dbContext.Orders
                                      join detail in dbContext.Order_Details on order.OrderID equals detail.OrderID
                                      join product in dbContext.Products on detail.ProductID equals product.ProductID
                                      where order.OrderDate.HasValue
                                      group new { product, detail } by new
                                      {
                                          Year = order.OrderDate.Value.Year,
                                          Month = order.OrderDate.Value.Month
                                      } into yearMonthGroup
                                      select new
                                      {
                                          Year = yearMonthGroup.Key.Year,
                                          Month = yearMonthGroup.Key.Month,
                                          TotalSales = yearMonthGroup.Sum(g => g.product.UnitPrice * g.detail.Quantity),
                                          Products = yearMonthGroup.Select(g => g.product).ToList()
                                      };
            // ---- 顯示到 DataGridView ----
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("年份", typeof(int));
            dataTable.Columns.Add("月份", typeof(int));
            dataTable.Columns.Add("總銷售金額", typeof(decimal));
            dataTable.Columns.Add("產品名稱", typeof(string));
            dataTable.Columns.Add("單價", typeof(decimal));

            foreach (var yearMonthGroup in productsByYearMonth)
            {
                foreach (var product in yearMonthGroup.Products)
                {
                    dataTable.Rows.Add(
                        yearMonthGroup.Year,
                        yearMonthGroup.Month,
                        yearMonthGroup.TotalSales,
                        product.ProductName,
                        product.UnitPrice
                    );
                }
            }

            // 綁定 DataTable 到 DataGridView
            dataGridView1.DataSource = dataTable;
            // ---- 顯示到 TreeView ----
            treeView1.Nodes.Clear(); // 清空之前的節點

            foreach (var yearMonthGroup in productsByYearMonth)
            {
                // 創建年份月份根節點，並顯示總銷售金額
                TreeNode yearMonthNode = new TreeNode($"年份: {yearMonthGroup.Year}, 月份: {yearMonthGroup.Month} - 總銷售金額: {yearMonthGroup.TotalSales:C2}");

                foreach (var product in yearMonthGroup.Products)
                {
                    // 每個產品作為子節點
                    TreeNode productNode = new TreeNode($"{product.ProductName} - {product.UnitPrice:C2}");
                    yearMonthNode.Nodes.Add(productNode);
                }

                // 將年份月份節點添加到 TreeView
                treeView1.Nodes.Add(yearMonthNode);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            // 根據 EmployeeID 查找銷售最好的 5 位業務員
            var topSalespeople = from order in dbContext.Orders
                                 join detail in dbContext.Order_Details on order.OrderID equals detail.OrderID
                                 join employee in dbContext.Employees on order.EmployeeID equals employee.EmployeeID
                                 where order.OrderDate.HasValue
                                 group new { detail, employee } by employee.EmployeeID into salesGroup
                                 let totalSales = salesGroup.Sum(g => g.detail.UnitPrice * g.detail.Quantity)
                                 orderby totalSales descending
                                 select new
                                 {
                                     Employee = salesGroup.FirstOrDefault().employee, // 使用 FirstOrDefault
                                     TotalSales = totalSales
                                 };

            // 取前 5 位銷售最好的業務員
            var top5Salespeople = topSalespeople.Take(5).ToList();

            // ---- 顯示到 DataGridView ----
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("業務員姓名", typeof(string));
            dataTable.Columns.Add("總銷售金額", typeof(decimal));

            foreach (var salesperson in top5Salespeople)
            {
                if (salesperson.Employee != null) // 確認 Employee 不為 null
                {
                    dataTable.Rows.Add(
                        salesperson.Employee.FirstName + " " + salesperson.Employee.LastName, // 顯示全名
                        salesperson.TotalSales
                    );
                }
            }

            // 綁定 DataTable 到 DataGridView
            dataGridView1.DataSource = dataTable;

            // ---- 顯示到 TreeView ----
            treeView1.Nodes.Clear(); // 清空之前的節點

            foreach (var salesperson in top5Salespeople)
            {
                if (salesperson.Employee != null) // 確認 Employee 不為 null
                {
                    // 創建業務員節點，並顯示總銷售金額
                    TreeNode salespersonNode = new TreeNode($"{salesperson.Employee.FirstName} {salesperson.Employee.LastName} - 總銷售金額: {salesperson.TotalSales:C2}");

                    // 將業務員節點添加到 TreeView
                    treeView1.Nodes.Add(salespersonNode);
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            // 查詢產品最高單價前 5 筆，包括類別名稱
            var topPricedProducts = (from product in dbContext.Products
                                     join category in dbContext.Categories on product.CategoryID equals category.CategoryID
                                     orderby product.UnitPrice descending
                                     select new
                                     {
                                         ProductName = product.ProductName,
                                         CategoryName = category.CategoryName,
                                         UnitPrice = product.UnitPrice
                                     }).Take(5).ToList();

            // ---- 顯示到 DataGridView ----
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("產品名稱", typeof(string));
            dataTable.Columns.Add("類別名稱", typeof(string));
            dataTable.Columns.Add("單價", typeof(decimal));

            foreach (var product in topPricedProducts)
            {
                dataTable.Rows.Add(
                    product.ProductName,
                    product.CategoryName,
                    product.UnitPrice
                );
            }

            // 綁定 DataTable 到 DataGridView
            dataGridView1.DataSource = dataTable;

            // ---- 顯示到 TreeView ----
            treeView1.Nodes.Clear(); // 清空之前的節點

            foreach (var product in topPricedProducts)
            {
                // 創建產品節點，並顯示類別名稱與單價
                TreeNode productNode = new TreeNode($"{product.ProductName} ({product.CategoryName}) - 單價: {product.UnitPrice:C2}");

                // 將產品節點添加到 TreeView
                treeView1.Nodes.Add(productNode);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            // 檢查是否有任何產品的單價大於 300
            bool hasExpensiveProduct = dbContext.Products.Any(p => p.UnitPrice > 300);

            if (hasExpensiveProduct)
            {
                // 取得所有單價大於 300 的產品資料
                var expensiveProducts = from p in dbContext.Products
                                        where p.UnitPrice > 300
                                        select new
                                        {
                                            產品名稱 = p.ProductName,
                                            單價 = p.UnitPrice
                                        };

                // 建立 DataTable 來顯示在 DataGridView
                DataTable dataTable = new DataTable();
                dataTable.Columns.Add("產品名稱", typeof(string));
                dataTable.Columns.Add("單價", typeof(decimal));

                // 將查詢結果添加到 DataTable
                foreach (var product in expensiveProducts)
                {
                    dataTable.Rows.Add(product.產品名稱, product.單價);
                }

                // 綁定 DataTable 到 DataGridView
                dataGridView1.DataSource = dataTable;
            }
            else
            {
                // 如果沒有單價大於 300 的產品，顯示 MessageBox
                MessageBox.Show("沒有任何產品的單價大於 300。");
            }
        }
    }
}
