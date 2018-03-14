using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using Mono.Data;
using Mono.Data.Sqlite;

namespace HeadlessL10nExporter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            label5.Hide();
        }

        private void OnBrowseClick(object sender, EventArgs e)
        {
            var openDialog = new OpenFileDialog()
            {
                Filter = "Files (*.xlsx)|*.xlsx"
            };
            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openDialog.FileName;
            }
        }

        private void OnExportClick(object sender, EventArgs e)
        {
            var xlsxPath = textBox1.Text;
            if (!Check(!string.IsNullOrWhiteSpace(xlsxPath), "Xlsx path is empty"))
            {
                return;
            }
            var fileInfo = new FileInfo(xlsxPath);
            if (!Check(fileInfo.Exists, "File does not exists"))
            {
                return;
            }
            ExcelPackage pack;
            string errMsg = "";
            try
            {
                pack = new ExcelPackage(fileInfo);
            }
            catch (Exception exc)
            {
                pack = null;
                errMsg = exc.Message;
            }
            if (!Check(pack != null, $"Cannot open xlsx document({errMsg})"))
            {
                return;
            }
            var dbPath = "";
            try
            {
                var metaSheet = pack.Workbook.Worksheets["meta"];
                var mainSheet = pack.Workbook.Worksheets["main"];
                if (metaSheet == null || mainSheet == null)
                {
                    throw new Exception("Cannot find 'meta' and 'main' sheet");
                }
                var showName = metaSheet.Cells[2, 1].Value?.ToString();
                var author = metaSheet.Cells[2, 2].Value?.ToString();
                var abbr = metaSheet.Cells[2, 3].Value?.ToString();
                if (string.IsNullOrWhiteSpace(showName) ||
                    string.IsNullOrWhiteSpace(author) ||
                    string.IsNullOrWhiteSpace(abbr))
                {
                    throw new Exception("Meta info is not filled correctly");
                }
                dbPath = showName.ToLower() + "_" + author.ToLower() + ".db";
                dbPath = Path.GetDirectoryName(xlsxPath) + "\\" + dbPath;
                label5.Show();
                label5.Refresh();
                var done = DoExport(mainSheet, dbPath, showName, author, abbr);
                label5.Hide();
                if (!done)
                {
                    // 用户取消
                    return;
                }
            }
            catch (Exception exc)
            {
                Check(false, "Export failed, " + exc.Message);
                label5.Hide();
                return;
            }

            pack.Dispose();
            MessageBox.Show($"Export done! Saved to '{dbPath}'", "Yeah!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            MessageBox.Show("Please manually copy the .db file to the language pack folder!", "Last Step", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private bool Check(bool condition, string message)
        {
            if (!condition)
            {
                MessageBox.Show(message + ", please check", "Oops...", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return condition;
        }

        private bool DoExport(ExcelWorksheet mainSheet, string dbPath, string showName, string author, string abbr)
        {
            if (File.Exists(dbPath))
            {
                // 文件已存在
                var result = MessageBox.Show("Database file is already exist! Overwrite?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (result == DialogResult.No)
                {
                    // 用户选择放弃
                    return false;
                }
                File.Delete(dbPath);
            }
            SqliteHelper helper = new SqliteHelper(dbPath);
            helper.Query("CREATE TABLE `meta` (`ShowName` TEXT, `Author` TEXT, `Abbr` TEXT, `Date` TEXT)");
            var exportDate = DateTime.UtcNow.ToString("yyyyMMdd");
            helper.Query($"INSERT INTO `meta` VALUES ('{showName}', '{author}', '{abbr}', '{exportDate}')");
            helper.Query("CREATE TABLE `main` (`Id` INT, `Alias` TEXT, `Text` TEXT)");
            int rowNo = 1;
            StringBuilder sqlString = new StringBuilder();
            while (true)
            {
                rowNo += 1;
                string id = mainSheet.Cells[rowNo, 1].Value?.ToString();
                string alias = mainSheet.Cells[rowNo, 2].Value?.ToString();
                string text = mainSheet.Cells[rowNo, 3].Value?.ToString();
                if (string.IsNullOrWhiteSpace(id))
                {
                    // 没有了
                    break;
                }
                if (string.IsNullOrWhiteSpace(alias))
                {
                    // alias为空
                    continue;
                }
                if (string.IsNullOrWhiteSpace(text))
                {
                    text = "";
                }
                Console.WriteLine("rowNo: " + rowNo);
                text = text.Replace("'", "''"); // 引号转义
                sqlString.AppendFormat("INSERT INTO `main` VALUES ({0}, '{1}', '{2}');", id, alias, text);
            }
            helper.Query(sqlString.ToString());
            helper.Dispose();
            return true;
        }
    }
}
