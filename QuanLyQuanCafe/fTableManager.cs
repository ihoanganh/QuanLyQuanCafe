using QuanLyQuanCafe.DAO;
using QuanLyQuanCafe.DTO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Button = System.Windows.Forms.Button;
using ComboBox = System.Windows.Forms.ComboBox;
using Excel = Microsoft.Office.Interop.Excel;

namespace QuanLyQuanCafe
{
    public partial class fTableManager : Form
    {
        private Account loginAccount;

        public Account LoginAccount
        {
            get { return loginAccount; }
            set { loginAccount = value; ChangeAccount(loginAccount.Type); }
        }
        public fTableManager(Account acc)
        {
            InitializeComponent();

            this.LoginAccount = acc;

            LoadTable();
            LoadCategory();
            LoadComboboxTable(cbSwitchTable);
            this.Icon = new Icon(Application.StartupPath + "\\coffee.ico");
        }

        #region Method

        void ChangeAccount(int type)
        {
            adminToolStripMenuItem.Enabled = (type == 1);
            thôngTinTàiKhoảnToolStripMenuItem.Text += " (" + LoginAccount.DisplayName + ")";
        }
        void LoadCategory()
        {
            List<Category> listCategory = CategoryDAO.Instance.GetListCategory();
            cbCategory.DataSource = listCategory;
            cbCategory.DisplayMember = "Name";
        }

        void LoadFoodListByCategoryID(int id)
        {
            List<Food> listFood = new List<Food>();
            listFood=  FoodDAO.Instance.GetFoodByCategoryID(id);
            cbFood.DataSource = listFood;
            cbFood.DisplayMember = "Name";
        }
        void LoadTable()
        {
            flpTable.Controls.Clear();
            
            List<Table> tableList = TableDAO.Instance.LoadTableList();

            foreach (Table item in tableList)
            {
                Button btn = new Button() { Width = TableDAO.TableWidth, Height = TableDAO.TableHeight};
                btn.Text = item.Name + Environment.NewLine + item.Status;
                btn.Click += btn_Click;
                btn.Tag = item;

                switch (item.Status)
                {
                    case "Trống":
                        btn.BackColor = Color.Aqua;
                        break;
                    default:
                        btn.BackColor = Color.Orange;
                        break;
                }

                flpTable.Controls.Add(btn);
            }
        }

        void ShowBill(int id)
        {
            lsvBill.Items.Clear();
            List<QuanLyQuanCafe.DTO.Menu> listBillInfo = MenuDAO.Instance.GetListMenuByTable(id);
            float totalPrice = 0;
            foreach (QuanLyQuanCafe.DTO.Menu item in listBillInfo)
            {
                ListViewItem lsvItem = new ListViewItem(item.FoodName.ToString());
                lsvItem.SubItems.Add(item.Count.ToString());
                lsvItem.SubItems.Add(item.Price.ToString());
                lsvItem.SubItems.Add(item.TotalPrice.ToString());
                totalPrice += item.TotalPrice;
                lsvBill.Items.Add(lsvItem);
            }
            CultureInfo culture = new CultureInfo("vi-VN");

            //Thread.CurrentThread.CurrentCulture = culture;

            txbTotalPrice.Text = totalPrice.ToString("c", culture);

        }

        void LoadComboboxTable(ComboBox cb)
        {
            cb.DataSource = TableDAO.Instance.LoadTableList();
            cb.DisplayMember = "Name";
        }
        void ResetValue()
        {
            cbCategory.Text = "";
            cbFood.Text = "";
        }
        #endregion


        #region Events

        private void thanhToánToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnCheckOut_Click(this, new EventArgs());
        }

        private void thêmMónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnAddFood_Click(this, new EventArgs());
        }
        private void xóaMónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Table table = lsvBill.Tag as Table;
            if (table == null)
            {
                MessageBox.Show("Hãy chọn bàn");
                return;
            }
            if (lsvBill.SelectedItems.Count > 0)
            {
                string tenmon = lsvBill.SelectedItems[0].SubItems[0].Text; 
                DialogResult dl = MessageBox.Show("Bạn muốn xóa", "canh bao", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dl == DialogResult.OK)
                {
                    string query = "Delete BillInfo where id = " +
                        "(select bi.id from BillInfo bi inner join Bill b on bi.idBill=b.id inner join Food f on bi.idFood =f.id " +
                        "where b.idTable="+table.ID+" and f.name=N'"+tenmon+"') ";
                    DataProvider.Instance.ExecuteQuery(query);
                }
                ShowBill(table.ID);
                LoadTable();

            }
            else MessageBox.Show("Hãy chọn món muốn xóa");
        }

        void btn_Click(object sender, EventArgs e)
        {
            int tableID = ((sender as Button).Tag as Table).ID;
            lsvBill.Tag = (sender as Button).Tag;
            ShowBill(tableID);
        }
        private void đăngXuấtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void thôngTinCáNhânToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fAccountProfile f = new fAccountProfile(LoginAccount);
            f.UpdateAccount += f_UpdateAccount;
            f.ShowDialog();
        }

        void f_UpdateAccount(object sender, AccountEvent e)
        {
            thôngTinTàiKhoảnToolStripMenuItem.Text = "Thông tin tài khoản (" + e.Acc.DisplayName + ")";
        }

        private void adminToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fAdmin f = new fAdmin();
            f.loginAccount = LoginAccount;
            f.InsertFood += f_InsertFood;
            f.DeleteFood += f_DeleteFood;
            f.UpdateFood += f_UpdateFood;
            f.InsertCategory += f_InsertCategory;
            f.DeleteCategory += f_DeleteCategory;
            f.UpdateCategory += f_UpdateCategory;
            f.InsertTable += F_InsertTable;
            f.DeleteTable += F_DeleteTable;
            f.UpdateTable += F_UpdateTable;
            f.ShowDialog();
        }

        private void F_UpdateTable(object sender, EventArgs e)
        {
            LoadTable();
        }

        private void F_DeleteTable(object sender, EventArgs e)
        {
            LoadTable();
        }

        private void F_InsertTable(object sender, EventArgs e)
        {
            LoadTable();
        }

        void f_UpdateFood(object sender, EventArgs e)
        {
            LoadFoodListByCategoryID((cbCategory.SelectedItem as Category).ID);
            if (lsvBill.Tag != null)
                ShowBill((lsvBill.Tag as Table).ID);
            ResetValue();
        }

        void f_DeleteFood(object sender, EventArgs e)
        {
            LoadFoodListByCategoryID((cbCategory.SelectedItem as Category).ID);
            if (lsvBill.Tag != null)
                ShowBill((lsvBill.Tag as Table).ID);
            ResetValue();
            LoadTable();
        }

        void f_InsertFood(object sender, EventArgs e)
        {
            LoadFoodListByCategoryID((cbCategory.SelectedItem as Category).ID);
            if (lsvBill.Tag != null)
                ShowBill((lsvBill.Tag as Table).ID);
            ResetValue();
        }
        void f_UpdateCategory(object sender, EventArgs e)
        {
            LoadCategory();
            if (lsvBill.Tag != null)
                ShowBill((lsvBill.Tag as Table).ID);
        }

        void f_DeleteCategory(object sender, EventArgs e)
        {
            LoadCategory();
            if (lsvBill.Tag != null)
                ShowBill((lsvBill.Tag as Table).ID);
            LoadTable();
        }

        void f_InsertCategory(object sender, EventArgs e)
        {
            LoadCategory();
            if (lsvBill.Tag != null)
                ShowBill((lsvBill.Tag as Table).ID);
        }

        private void cbCategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            int id = 0;

            ComboBox cb = sender as ComboBox;

            if (cb.SelectedItem == null)
                return;

            Category selected = cb.SelectedItem as Category;
            id = selected.ID;
            cbFood.Text = "";
            LoadFoodListByCategoryID(id);
        }

        private void btnAddFood_Click(object sender, EventArgs e)
        {
            Table table = lsvBill.Tag as Table;

            if (table == null)
            {
                MessageBox.Show("Hãy chọn bàn");
                return;
            }

            int idBill = BillDAO.Instance.GetUncheckBillIDByTableID(table.ID);
            if (cbFood.Text == "")
            {
                MessageBox.Show("Món không tồn tại.Hãy chọn lại");
                return;
            }
            int  foodID = (cbFood.SelectedItem as Food).ID;
            
            int count = (int)nmFoodCount.Value;

            if (idBill == -1)
            {
                BillDAO.Instance.InsertBill(table.ID);
                BillInfoDAO.Instance.InsertBillInfo(BillDAO.Instance.GetMaxIDBill(), foodID , count);
            }
            else
            {
                BillInfoDAO.Instance.InsertBillInfo(idBill, foodID, count);
            }

            ShowBill(table.ID);

            LoadTable();
        }        
        private void btnCheckOut_Click(object sender, EventArgs e)
        {
            Table table = lsvBill.Tag as Table;

            int idBill = BillDAO.Instance.GetUncheckBillIDByTableID(table.ID);
            int discount = (int)nmDisCount.Value;

            double totalPrice = Convert.ToDouble(txbTotalPrice.Text.Split(',')[0]);
            double finalTotalPrice = totalPrice - (totalPrice/100)*discount;

            if (idBill != -1)
            {
                if (MessageBox.Show(string.Format("Bạn có chắc thanh toán hóa đơn cho bàn {0}\nTổng tiền - (Tổng tiền / 100) x Giảm giá\n=> {1} - ({1} / 100) x {2} = {3}",table.Name, totalPrice, discount, finalTotalPrice), "Thông báo", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
                {
                    BillDAO.Instance.CheckOut(idBill, discount, (float)finalTotalPrice);
                    //xuất hóa đơn ra file excel
                    Excel.Application exApp = new Excel.Application();
                    Excel.Workbook exBook = exApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    Excel.Worksheet exSheet = (Excel.Worksheet)exBook.Worksheets[1];
                    Excel.Range exRange = (Excel.Range)exSheet.Cells[1, 1]; //Đưa con trỏ vào ô A1
                    
                    exSheet.Range["A1:F1"].MergeCells = true;
                    exSheet.Range["A1:F1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    exSheet.Range["A1:F1"].Font.Size = 15;
                    exSheet.Range["A1:F1"].Font.Bold = true;
                    exSheet.Range["A1:F1"].Font.Color = Color.Blue;
                    exSheet.Range["A1:F1"].Value = "AI COFFEE";

                    Excel.Range dc = exSheet.Range["A2:F2"];
                    dc.MergeCells = true;
                    dc.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    dc.Font.Size = 10;
                    dc.Font.Bold = true;
                    dc.Font.Color = Color.Red;
                    dc.Value = "Địa chỉ: Trường Đại Học Giao Thông Vận Tải Hà Nội";
                    //in chu hoa don ban
                    exSheet.Range["A4:F4"].MergeCells = true;
                    exSheet.Range["A4:F4"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    exSheet.Range["A4:F4"].Font.Size = 18; //dua con tro vao o A1
                    exSheet.Range["A4:F4"].Font.Bold = true;
                    exSheet.Range["A4:F4"].Font.Color = Color.Red;
                    exSheet.Range["A4:F4"].Value = "HOÁ ĐƠN";

                    //in các thông tin chung
                    exSheet.Range["A6"].Font.Size = 11;
                    exSheet.Range["A6"].Value = "Bàn " + table.ID;
                    exSheet.Range["A7"].Value = "Mã hóa đơn: "+idBill;

                    exSheet.Range["A8:E8"].Font.Size = 12;
                    exSheet.Range["A8"].Value = "STT ";
                    exSheet.Range["B8"].Value = "Món";
                    exSheet.Range["C8"].Value = "Số lượng:";
                    exSheet.Range["D8"].Value = "Đơn giá";
                    exSheet.Range["E8"].Value = "Thành tiền";
                    int i = 9;
                    int j = 1;

                    foreach (ListViewItem item in lsvBill.Items)
                    {
                        exSheet.Cells[i, 1] = j;
                        exSheet.Cells[i, 2] = item.SubItems[0].Text;
                        exSheet.Cells[i, 3] = item.SubItems[1].Text;
                        exSheet.Cells[i, 4] = item.SubItems[2].Text;
                        exSheet.Cells[i, 5] = item.SubItems[3].Text;
                        i++;
                        j++;

                    }
                    exSheet.Cells[i + 2, 4] = "Giảm giá: ";
                    exSheet.Cells[i + 2, 5] = nmDisCount.Text;
                    exSheet.Cells[i + 2, 6] = "(%)";
                    exSheet.Cells[i + 3, 4] = "Tổng tiền: ";
                    exSheet.Cells[i + 3, 5] = finalTotalPrice;
                    exSheet.Cells[i + 3, 6] = "đồng";
                    //in dòng tiêu đề



                    //in

                    exBook.Activate();

                    //lưu file
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx|All File|*.*";
                    save.FilterIndex = 2;
                    save.FileName = "hoadon" +idBill;
                    if (save.ShowDialog() == DialogResult.OK)
                    {
                        exBook.SaveAs(save.FileName.ToLower());
                    }
                    exApp.Quit();

                    ShowBill(table.ID);

                    LoadTable();
                }
            }
        }
        private void btnSwitchTable_Click(object sender, EventArgs e)
        {           

            int id1 = (lsvBill.Tag as Table).ID;

            int id2 = (cbSwitchTable.SelectedItem as Table).ID;
            if (MessageBox.Show(string.Format("Bạn có thật sự muốn chuyển bàn {0} qua bàn {1}", (lsvBill.Tag as Table).Name, (cbSwitchTable.SelectedItem as Table).Name), "Thông báo", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
            {
                TableDAO.Instance.SwitchTable(id1, id2);

                LoadTable();
            }
        }


        #endregion

        
    }
}
