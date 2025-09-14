using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Text;

namespace FinCompare
{
    public partial class Form1 : Form
    {
        private DataTable dataTable = new();
        private bool isCalculating = false; // 防止递归计算

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            InitializeDataTable();
            SetupDataGridView();
            AddSampleRow();
        }

        private void InitializeDataTable()
        {
            dataTable = new DataTable();
            dataTable.Columns.Add("名称/机构", typeof(string));
            dataTable.Columns.Add("借款/分期金额", typeof(decimal));
            dataTable.Columns.Add("期数", typeof(int));
            dataTable.Columns.Add("平均每期利息", typeof(decimal));
            dataTable.Columns.Add("总利息", typeof(decimal));
            dataTable.Columns.Add("利息率", typeof(string));
            dataTable.Columns.Add("近似折算年化利率（单利）", typeof(string));
            dataTable.Columns.Add("添加时间", typeof(DateTime));

            dataGridView1.DataSource = dataTable;
        }

        private void SetupDataGridView()
        {
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.AllowUserToDeleteRows = true;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            
            // 启用排序功能
            dataGridView1.AllowUserToOrderColumns = true;
            
            // 设置列宽比例
            dataGridView1.Columns[0].FillWeight = 15; // 名称/机构
            dataGridView1.Columns[1].FillWeight = 15; // 借款/分期金额
            dataGridView1.Columns[2].FillWeight = 10; // 期数
            dataGridView1.Columns[3].FillWeight = 15; // 平均每期利息
            dataGridView1.Columns[4].FillWeight = 15; // 总利息
            dataGridView1.Columns[5].FillWeight = 10; // 利息率
            dataGridView1.Columns[6].FillWeight = 15; // 近似折算年化利率
            dataGridView1.Columns[7].FillWeight = 15; // 添加时间
        
            // 设置数值列的格式
            dataGridView1.Columns[1].DefaultCellStyle.Format = "N2";
            dataGridView1.Columns[3].DefaultCellStyle.Format = "N2";
            dataGridView1.Columns[4].DefaultCellStyle.Format = "N2";
            dataGridView1.Columns[7].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";
        
            // 设置列的排序模式
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.Programmatic;
            }
        }

        private void AddSampleRow()
        {
            DataRow row = dataTable.NewRow();
            row["名称/机构"] = "";
            row["借款/分期金额"] = 0m;
            row["期数"] = 0;
            row["平均每期利息"] = 0m;
            row["总利息"] = 0m;
            row["利息率"] = "";
            row["近似折算年化利率（单利）"] = "";
            row["添加时间"] = DateTime.Now;
            dataTable.Rows.Add(row);
        }

        private void DataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (isCalculating || e.RowIndex < 0) return;

            DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
            
            // 如果是新行且添加时间为空，设置添加时间
            if (row.Cells["添加时间"].Value == null || row.Cells["添加时间"].Value == DBNull.Value)
            {
                row.Cells["添加时间"].Value = DateTime.Now;
            }

            CalculateFinancialMetrics(e.RowIndex);
        }

        private void CalculateFinancialMetrics(int rowIndex)
        {
            if (isCalculating) return;
            isCalculating = true;

            try
            {
                DataGridViewRow row = dataGridView1.Rows[rowIndex];
                
                // 获取基础数据并确保数值字段不为空
                decimal principal = GetDecimalValue(row.Cells["借款/分期金额"].Value);
                int periods = GetIntValue(row.Cells["期数"].Value);
                decimal monthlyInterest = GetDecimalValue(row.Cells["平均每期利息"].Value);
                decimal totalInterest = GetDecimalValue(row.Cells["总利息"].Value);

                // 确保数值字段显示为0而不是空值
                row.Cells["借款/分期金额"].Value = principal;
                row.Cells["期数"].Value = periods;
                row.Cells["平均每期利息"].Value = monthlyInterest;
                row.Cells["总利息"].Value = totalInterest;

                if (principal <= 0 || periods <= 0) return;

                // 根据输入情况进行计算
                if (monthlyInterest > 0 && totalInterest == 0)
                {
                    // 有每期利息，计算总利息
                    totalInterest = monthlyInterest * periods;
                    row.Cells["总利息"].Value = totalInterest;
                }
                else if (totalInterest > 0 && monthlyInterest == 0)
                {
                    // 有总利息，计算每期利息
                    monthlyInterest = totalInterest / periods;
                    row.Cells["平均每期利息"].Value = monthlyInterest;
                }

                if (totalInterest > 0)
                {
                    // 计算利息率
                    decimal interestRate = (totalInterest / principal) * 100;
                    row.Cells["利息率"].Value = $"{interestRate:F2}%";

                    // 计算年化利率
                    double annualRate = CalculateAnnualRate(principal, periods, monthlyInterest);
                    row.Cells["近似折算年化利率（单利）"].Value = $"{annualRate:F2}%";
                }
            }
            finally
            {
                isCalculating = false;
            }
        }

        private static double CalculateAnnualRate(decimal principal, int periods, decimal monthlyInterest)
        {
            double principalAmount = (double)principal;
            double monthlyPayment = (double)(principal / periods + monthlyInterest);
            
            // 使用牛顿迭代法计算IRR
            double irr = 0.01; // 初始值1%
            const double tolerance = 0.000001;
            const int maxIterations = 100;

            for (int i = 0; i < maxIterations; i++)
            {
                double npv = -principalAmount;
                double npvDerivative = 0;

                for (int period = 1; period <= periods; period++)
                {
                    double factor = Math.Pow(1 + irr, period);
                    npv += monthlyPayment / factor;
                    npvDerivative -= period * monthlyPayment / Math.Pow(1 + irr, period + 1);
                }

                if (Math.Abs(npv) < tolerance) break;

                double newIrr = irr - npv / npvDerivative;
                if (Math.Abs(newIrr - irr) < tolerance) break;
                
                irr = newIrr;
            }

            return irr * 12 * 100; // 转换为年化百分比
        }

        private static decimal GetDecimalValue(object? value)
        {
            if (value == null || value == DBNull.Value || string.IsNullOrEmpty(value.ToString())) return 0;
            if (decimal.TryParse(value.ToString(), out decimal result)) return result;
            return 0;
        }

        private static int GetIntValue(object? value)
        {
            if (value == null || value == DBNull.Value || string.IsNullOrEmpty(value.ToString())) return 0;
            if (int.TryParse(value.ToString(), out int result)) return result;
            return 0;
        }

        private void DataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                DataGridViewColumn column = dataGridView1.Columns[e.ColumnIndex];
                
                // 确定排序方向
                ListSortDirection direction;
                if (dataGridView1.SortedColumn == column && dataGridView1.SortOrder == SortOrder.Ascending)
                {
                    direction = ListSortDirection.Descending;
                }
                else
                {
                    direction = ListSortDirection.Ascending;
                }
        
                // 使用DataView排序
                string columnName = dataTable.Columns[e.ColumnIndex].ColumnName;
                dataTable.DefaultView.Sort = $"[{columnName}] {(direction == ListSortDirection.Ascending ? "ASC" : "DESC")}";
        
                // 更新DataGridView排序状态
                dataGridView1.Sort(column, direction);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"排序时发生错误：{ex.Message}", "排序错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void 保存为CSVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDialog = new()
            {
                Filter = "CSV文件|*.csv",
                Title = "保存为CSV文件"
            };

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                SaveToCsv(saveDialog.FileName);
            }
        }

        private void SaveToCsv(string fileName)
        {
            try
            {
                StringBuilder csv = new();
                
                // 添加列标题
                string[] headers = new string[dataTable.Columns.Count];
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    headers[i] = dataTable.Columns[i].ColumnName;
                }
                csv.AppendLine(string.Join(",", headers.Select(EscapeCsvField)));

                // 添加数据行
                foreach (DataRow row in dataTable.Rows)
                {
                    string[] fields = new string[dataTable.Columns.Count];
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        fields[j] = row[j]?.ToString() ?? "";
                    }
                    csv.AppendLine(string.Join(",", fields.Select(EscapeCsvField)));
                }

                File.WriteAllText(fileName, csv.ToString(), Encoding.UTF8);
                MessageBox.Show("CSV文件保存成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存CSV文件失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string EscapeCsvField(string field)
        {
            if (field.Contains(',') || field.Contains('"') || field.Contains('\n'))
            {
                return "\"" + field.Replace("\"", "\"\"") + "\"";
            }
            return field;
        }

        private void 加载CSVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new()
            {
                Filter = "CSV文件|*.csv",
                Title = "加载CSV文件"
            };

            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                LoadFromCsv(openDialog.FileName);
            }
        }

        private void LoadFromCsv(string fileName)
        {
            try
            {
                dataTable.Clear();
                string[] lines = File.ReadAllLines(fileName, Encoding.UTF8);
                
                if (lines.Length > 1)
                {
                    for (int i = 1; i < lines.Length; i++)
                    {
                        string[] fields = ParseCsvLine(lines[i]);
                        if (fields.Length >= dataTable.Columns.Count)
                        {
                            DataRow row = dataTable.NewRow();
                            for (int j = 0; j < dataTable.Columns.Count; j++)
                            {
                                if (j < fields.Length)
                                {
                                    object value = ConvertToColumnType(fields[j], dataTable.Columns[j].DataType);
                                    row[j] = value;
                                }
                            }
                            dataTable.Rows.Add(row);
                        }
                    }
                }

                MessageBox.Show("CSV文件加载成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载CSV文件失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string[] ParseCsvLine(string line)
        {
            List<string> fields = [];
            bool inQuotes = false;
            StringBuilder currentField = new();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                
                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        currentField.Append('"');
                        i++; // 跳过下一个引号
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    fields.Add(currentField.ToString());
                    currentField.Clear();
                }
                else
                {
                    currentField.Append(c);
                }
            }
            
            fields.Add(currentField.ToString());
            return [.. fields];
        }

        private static object ConvertToColumnType(string value, Type targetType)
        {
            if (string.IsNullOrEmpty(value)) return DBNull.Value;

            try
            {
                if (targetType == typeof(decimal))
                {
                    return decimal.Parse(value);
                }
                else if (targetType == typeof(int))
                {
                    return int.Parse(value);
                }
                else if (targetType == typeof(DateTime))
                {
                    return DateTime.Parse(value);
                }
                else
                {
                    return value;
                }
            }
            catch
            {
                return DBNull.Value;
            }
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
