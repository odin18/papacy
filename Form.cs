using Seagull.BarTender.Print;
using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Printing;
using System.Management;
using System.Collections.Generic;

namespace papacy1
{
    public partial class papacy : Form
    {
        private Engine engine;
        private LabelFormatDocument btFormat = null;
        private double currentPrintCount;
        private string tempPath;
        private string selectedPrinter;
        private int copies;
        private int SameCNOcopies;
        private bool isUpdating = false;
        private List<TabPage> _allTabControl = new List<TabPage>();


        //出bug解決方法 https://www.796t.com/content/1547133874.html
        public papacy()
        {

            InitializeComponent();

            copies = 0;

            // 設置DrawMode為OwnerDrawFixed允許自定義繪圖
            tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;

            // 訂閱DrawItem事件
            tabControl.DrawItem += new DrawItemEventHandler(tabControl_DrawItem);

            //Properties.Settings.Default.Reset();

            // 添加已安裝的印表機到ComboBox
            LoadPrinters();
            // 設置ComboBox高度，高度等於每個項目的高度乘以3
            listBoxPrinters.Height = 3 * listBoxPrinters.ItemHeight;

            // 如果沒有預設的印表機選擇
            if (Properties.Settings.Default.SelectedPrinter == null)
            {
                PrintMachinelabel.Text = "尚未選擇";
            }
            // 如果預設的印表機不在已安裝的印表機列表中
            else if (!listBoxPrinters.Items.Contains(Properties.Settings.Default.SelectedPrinter))
            {
                PrintMachinelabel.Text = "尚未選擇";
                // 重置預設的印表機選擇
                Properties.Settings.Default.SelectedPrinter = null;
            }
            else
            {
                // 顯示預設選擇的印表機名稱
                PrintMachinelabel.Text = Properties.Settings.Default.SelectedPrinter;
                listBoxPrinters.SelectedItem = Properties.Settings.Default.SelectedPrinter;
            }
            PrintQuantitynumericUpDown.Value = Properties.Settings.Default.Copies;

            try
            {
                engine = new Engine();
            }
            catch (Exception)
            {
                MessageBox.Show("請安裝正確的BarTender");
            }

            MD1_CBX_CNO.Text = "3";
            MD2_CBX_CNO.Text = "4";
            MD3_CBX_CNO.Text = "3";
            MD4_CBX_CNO.Text = "3";
            MD5_CBX_CNO.Text = "3";
            MD6_CBX_CNO.Text = "3";
            MD7_CBX_CNO.Text = "3";

            foreach(TabPage page in tabControl.TabPages)
            {
                _allTabControl.Add(page);
            }

            tabControl.TabPages.Clear();

            // 將tabControl的當前選定標籤頁設定為tabPage8
            tabControl.TabPages.Add(_allTabControl.Where(x => x.Text == "列印設定").First());
        }
        private void LoadPrinters()
        {
            listBoxPrinters.Items.Clear(); // 清除ListBox中現有的所有項目

            // 使用WMI查詢來獲取所有Win32_Printer對象（也就是所有印表機）
            string query = "Select * From Win32_Printer";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);

            // 遍歷每一個查詢結果（每一個印表機）
            foreach (ManagementObject printer in searcher.Get())
            {
                // 獲取印表機的WorkOffline屬性，並判斷是否為離線狀態
                bool isPrinterOffline = printer["WorkOffline"].ToString().ToLower().Equals("true");

                // 只將在線的印表機添加到listBoxPrinters中
                if (!isPrinterOffline)
                {
                    listBoxPrinters.Items.Add(printer["Name"]);
                }
            }
        }
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 根據選定的索引（標籤）來顯示相應的tabPage
            switch (tabControl.SelectedIndex)
            {
                case 0:
                    tabPage1.Show();
                    break;
                case 1:
                    tabPage2.Show();
                    break;
                case 2:
                    tabPage3.Show();
                    break;
                case 3:
                    tabPage4.Show();
                    break;
                case 4:
                    tabPage5.Show();
                    break;
                case 5:
                    tabPage6.Show();
                    break;
                case 6:
                    tabPage7.Show();
                    break;
                case 7:
                    tabPage8.Show();
                    break;
                default:
                    break;
            }
        }

        private void tabControl_DrawItem(object sender, DrawItemEventArgs e)
        {
            // 獲取TabControl和TabPage的參考
            tabControl = sender as TabControl;
            TabPage tabPage = tabControl.TabPages[e.Index];
            Rectangle tabRect = tabControl.GetTabRect(e.Index);

            // 根據選項卡是否被選中來設定背景顏色
            Brush backBrush = (e.State == DrawItemState.Selected) ? new SolidBrush(Color.CornflowerBlue) : new SolidBrush(Color.Lavender);
            e.Graphics.FillRectangle(backBrush, tabRect);  // 繪制背景

            // 設定文字顏色和字體
            Brush foreBrush = new SolidBrush(Color.Black);
            Font RegularFont = new Font(e.Font, FontStyle.Regular);
            StringFormat stringFormat = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            e.Graphics.DrawString(tabPage.Text, RegularFont, foreBrush, tabRect, stringFormat);  // 繪制文字

            // 處理選項卡控件右側的空白區域
            if (e.Index == tabControl.TabCount - 1)
            {
                // 計算和繪制右側空白區域
                Rectangle rightRect = new Rectangle(tabRect.Right, tabRect.Top, tabControl.Width - tabRect.Right, tabRect.Height);
                Brush rightBrush = new SolidBrush(Color.Transparent);
                e.Graphics.FillRectangle(rightBrush, rightRect);  // 繪制右側空白區域
                rightBrush.Dispose();
            }

            backBrush.Dispose();
            foreBrush.Dispose();
            RegularFont.Dispose();
        }


        private void papacy_Resize(object sender, EventArgs e)
        {
            // 計算新的寬度和高度相對於原始尺寸(1028x768)的比例
            float newx = (this.Width) / 1028f;
            float newy = (this.Height) / 768f;

            // 根據新的尺寸比例調整TabControl的項目大小
            tabControl.ItemSize = new Size((int)(100 * newx), (int)(40 * newy));

            // 調用setControls方法來重新設置其他控件的尺寸和位置
            setControls(newx, newy, this);
        }

        private void setControls(float newx, float newy, Control cons)
        {
            //動態地調整控件尺寸和位置
            foreach (Control con in cons.Controls)
            {
                // 如果控件的Tag屬性不為null，則執行調整操作
                if (con.Tag != null)
                {
                    // 將Tag中的數據（以分號分隔）拆分為數組
                    string[] mytag = con.Tag.ToString().Split(new char[] { ';' });

                    // 根據新的尺寸比例調整控件的寬度、高度、左邊距和上邊距
                    con.Width = Convert.ToInt32(System.Convert.ToSingle(mytag[0]) * newx);
                    con.Height = Convert.ToInt32(System.Convert.ToSingle(mytag[1]) * newy);
                    con.Left = Convert.ToInt32(System.Convert.ToSingle(mytag[2]) * newx);
                    con.Top = Convert.ToInt32(System.Convert.ToSingle(mytag[3]) * newy);

                    // 根據新的高度比例調整字體大小
                    Single currentSize = System.Convert.ToSingle(mytag[4]) * newy;
                    con.Font = new Font(con.Font.Name, currentSize, con.Font.Style, con.Font.Unit);

                    // 如果該控件還有子控件，遞歸調用此方法
                    if (con.Controls.Count > 0)
                    {
                        setControls(newx, newy, con);
                    }
                }
            }
        }

        private void papacy_Load(object sender, EventArgs e)
        {
            // 設定 LOTtextBoxes 的初始值
            for (int i = 1; i <= 7; i++)
            {
                if (i != 3) // 跳過3
                {
                    string settingName = $"LOTNum{i}";
                    string propertyName = $"LOTtextBox{i}";
                    var settingValue = Properties.Settings.Default[settingName];
                    if (settingValue != null && !string.IsNullOrEmpty(settingValue.ToString()))
                    {
                        var textBox = this.Controls.Find(propertyName, true).FirstOrDefault() as TextBox;
                        if (textBox != null)
                        {
                            textBox.Text = settingValue.ToString();
                        }
                    }
                }
            }

            // 設定 OrigintextBox 和 EndtextBox
            for (int i = 0; i < tabControl.TabPages.Count; i++)
            {
                System.Windows.Forms.TextBox originTextBox = (System.Windows.Forms.TextBox)tabControl.TabPages[i].Controls["OrigintextBox" + (i + 1)];
                System.Windows.Forms.TextBox endTextBox = (System.Windows.Forms.TextBox)tabControl.TabPages[i].Controls["EndtextBox" + (i + 1)];

                if (originTextBox != null)
                {
                    originTextBox.Text = "MADE IN";
                    originTextBox.ForeColor = SystemColors.GrayText;
                    originTextBox.TextAlign = HorizontalAlignment.Left;
                }

                if (endTextBox != null)
                {
                    endTextBox.ForeColor = SystemColors.WindowText;
                }
            }

            // 設置每個控件的Tag屬性，用於後續的尺寸和位置調整
            SetTags(this);

        }

        private void SetTags(Control cons)
        {
            foreach (Control con in cons.Controls)
            {
                // 設置控件的Tag屬性，其中包括控件的寬度、高度、左邊距、上邊距和字體大小
                con.Tag = $"{con.Width};{con.Height};{con.Left};{con.Top};{con.Font.Size}";
                // 如果該控件還有子控件，遞歸地設置它們的Tag
                if (con.Controls.Count > 0)
                {
                    SetTags(con);
                }
            }
        }

        private void PrintQuantitynumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            // 將新的列印數量存儲到Properties.Settings.Default.Copies中
            Properties.Settings.Default.Copies = (int)PrintQuantitynumericUpDown.Value;
            // 儲存設定
            Properties.Settings.Default.Save();
        }

        private void Printerbutton_Click(object sender, EventArgs e)
        {
            // 檢查是否有選擇打印機
            if (listBoxPrinters.SelectedItem != null)
            {
                // 將選定的打印機名稱存儲到設定中
                Properties.Settings.Default.SelectedPrinter = listBoxPrinters.SelectedItem.ToString();
                // 更新顯示選定的打印機名稱的標籤
                PrintMachinelabel.Text = Properties.Settings.Default.SelectedPrinter;
                // 儲存設定
                Properties.Settings.Default.Save();
            }
        }
        private void ResearchButton_Click(object sender, EventArgs e)
        {
            // 清除選定的打印機設定
            Properties.Settings.Default.SelectedPrinter = null;
            // 如果沒有選定打印機，則顯示"尚未選擇"
            if (Properties.Settings.Default.SelectedPrinter == null)
            {
                PrintMachinelabel.Text = "尚未選擇";
            }
            // 儲存設定
            Properties.Settings.Default.Save();
            // 重新載入可用的打印機列表
            LoadPrinters();
        }

        private void DeleteDefaultOrigintextBox(int page)
        {
            // 根據指定的頁面和控件名稱找到OrigintextBox
            System.Windows.Forms.TextBox originTextBox = (System.Windows.Forms.TextBox)tabControl.TabPages[page - 1].Controls["OrigintextBox" + (page)];

            // 檢查TextBox中是否是預設的文字"MADE IN"
            if (originTextBox.Text == "MADE IN")
            {
                // 如果是預設文字，則清空TextBox
                originTextBox.Text = "";

                // 修改文字顏色為窗口的預設顏色
                originTextBox.ForeColor = SystemColors.WindowText;

                // 將文字對齊設為居中
                originTextBox.TextAlign = HorizontalAlignment.Center;
            }
        }

        private void OrigintextBox1_Enter(object sender, EventArgs e)
        {
            DeleteDefaultOrigintextBox(1);
        }

        private void OrigintextBox3_Enter(object sender, EventArgs e)
        {
            DeleteDefaultOrigintextBox(3);
        }

        private void OrigintextBox4_Enter(object sender, EventArgs e)
        {
            DeleteDefaultOrigintextBox(4);
        }

        private void OrigintextBox5_Enter(object sender, EventArgs e)
        {
            DeleteDefaultOrigintextBox(5);
        }

        private void OrigintextBox6_Enter(object sender, EventArgs e)
        {
            DeleteDefaultOrigintextBox(6);
        }

        private void OrigintextBox7_Enter(object sender, EventArgs e)
        {
            DeleteDefaultOrigintextBox(7);
        }
        private void setDefaultOrigintextBox(int page)
        {
            // 根據指定的頁面和控件名稱找到OrigintextBox
            System.Windows.Forms.TextBox originTextBox = (System.Windows.Forms.TextBox)tabControl.TabPages[page - 1].Controls["OrigintextBox" + (page)];

            // 檢查TextBox是否為空
            if (string.IsNullOrEmpty(originTextBox.Text))
            {
                // 如果是空的，則設置為預設文字"MADE IN"
                originTextBox.Text = "MADE IN";

                // 修改文字顏色為灰色
                originTextBox.ForeColor = SystemColors.GrayText;

                // 將文字對齊設為左對齊
                originTextBox.TextAlign = HorizontalAlignment.Left;
            }
        }

        private void OrigintextBox1_Leave(object sender, EventArgs e)
        {
            setDefaultOrigintextBox(1);
            // 移動到下一個控件
            SPECtextBox1.Focus();
        }

        private void OrigintextBox3_Leave(object sender, EventArgs e)
        {
            setDefaultOrigintextBox(3);
            // 移動到下一個控件
            CNOtextBox3.Focus();
        }

        private void OrigintextBox4_Leave(object sender, EventArgs e)
        {
            setDefaultOrigintextBox(4);
            // 移動到下一個控件
            CNOtextBox4.Focus();
        }

        private void OrigintextBox5_Leave(object sender, EventArgs e)
        {
            setDefaultOrigintextBox(5);
            // 移動到下一個控件
            SPECtextBox5.Focus();
        }

        private void OrigintextBox6_Leave(object sender, EventArgs e)
        {
            setDefaultOrigintextBox(6);
            // 移動到下一個控件
            MaterialtextBox6.Focus();
        }

        private void OrigintextBox7_Leave(object sender, EventArgs e)
        {
            setDefaultOrigintextBox(7);
            // 移動到下一個控件
            CartontextBox7.Focus();
        }

        private void OnlyAllowDigits(KeyPressEventArgs e)
        {
            // 只允许输入数字、退格键和删除键
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // 阻止输入事件的传播
            }
        }
        private void OnlyAllowDigitsAndOneDot(KeyPressEventArgs e, System.Windows.Forms.TextBox textBox)
        {
            // 只允許數字、控制鍵、小數點
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // 只允許一個小數點
            if ((e.KeyChar == '.') && (textBox.Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void NWtextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigitsAndOneDot(e, sender as System.Windows.Forms.TextBox);
        }

        private void GWtextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigitsAndOneDot(e, sender as System.Windows.Forms.TextBox);
        }

        private void StarttextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigits(e);
        }

        private void QuantitytextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigits(e);
        }

        private void NWtextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigitsAndOneDot(e, sender as System.Windows.Forms.TextBox);
        }

        private void GWtextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigitsAndOneDot(e, sender as System.Windows.Forms.TextBox);
        }

        private void StarttextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigits(e);
        }

        private void NWtextBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigitsAndOneDot(e, sender as System.Windows.Forms.TextBox);
        }

        private void GWtextBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigitsAndOneDot(e, sender as System.Windows.Forms.TextBox);
        }

        private void StarttextBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigits(e);
        }

        private void NWtextBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigitsAndOneDot(e, sender as System.Windows.Forms.TextBox);
        }

        private void GWtextBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigitsAndOneDot(e, sender as System.Windows.Forms.TextBox);
        }

        private void StarttextBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigits(e);
        }

        private void NWtextBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigitsAndOneDot(e, sender as System.Windows.Forms.TextBox);
        }

        private void GWtextBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigitsAndOneDot(e, sender as System.Windows.Forms.TextBox);
        }

        private void StarttextBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigits(e);
        }

        private void NWtextBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigitsAndOneDot(e, sender as System.Windows.Forms.TextBox);
        }

        private void GWtextBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigitsAndOneDot(e, sender as System.Windows.Forms.TextBox);
        }

        private void StarttextBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigits(e);
        }

        private void StarttextBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigits(e);
        }

        private void AddtextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigits(e);
        }

        private void AddtextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigits(e);
        }

        private void AddtextBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigits(e);
        }
        private void AddtextBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigits(e);
        }

        private void AddtextBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigits(e);
        }

        private void AddtextBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigits(e);
        }

        private void AddtextBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyAllowDigits(e);
        }

        private void GWComboBoxUnits(int page)
        {
            // 找到指定頁面中的GWunitcomboBox和NWunitcomboBox
            System.Windows.Forms.ComboBox GWunitcomboBox = (System.Windows.Forms.ComboBox)tabControl.TabPages[page - 1].Controls["GWunitcomboBox" + page];
            System.Windows.Forms.ComboBox NWunitcomboBox = (System.Windows.Forms.ComboBox)tabControl.TabPages[page - 1].Controls["NWunitcomboBox" + page];

            // 若GWunitcomboBox有選中項且NWunitcomboBox沒有選中項，則將GWunitcomboBox的選中項設為NWunitcomboBox的選中項
            if (GWunitcomboBox.SelectedIndex != -1 && NWunitcomboBox.SelectedIndex == -1)
            {
                NWunitcomboBox.SelectedItem = GWunitcomboBox.SelectedItem;
            }
        }
        private void NWComboBoxUnits(int page)
        {
            // 找到指定頁面中的GWunitcomboBox和NWunitcomboBox
            System.Windows.Forms.ComboBox GWunitcomboBox = (System.Windows.Forms.ComboBox)tabControl.TabPages[page - 1].Controls["GWunitcomboBox" + page];
            System.Windows.Forms.ComboBox NWunitcomboBox = (System.Windows.Forms.ComboBox)tabControl.TabPages[page - 1].Controls["NWunitcomboBox" + page];

            // 若NWunitcomboBox有選中項且GWunitcomboBox沒有選中項，則將NWunitcomboBox的選中項設為GWunitcomboBox的選中項
            if (NWunitcomboBox.SelectedIndex != -1 && GWunitcomboBox.SelectedIndex == -1)
            {
                GWunitcomboBox.SelectedItem = NWunitcomboBox.SelectedItem;
            }
        }

        private void NWunitcomboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            NWComboBoxUnits(1);
        }

        private void GWunitcomboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            GWComboBoxUnits(1);
        }

        private void NWunitcomboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            NWComboBoxUnits(2);
        }

        private void GWunitcomboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            GWComboBoxUnits(2);
        }

        private void NWunitcomboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            NWComboBoxUnits(3);
        }

        private void GWunitcomboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            GWComboBoxUnits(3);
        }

        private void NWunitcomboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            NWComboBoxUnits(4);
        }

        private void GWunitcomboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            GWComboBoxUnits(4);
        }

        private void NWunitcomboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            NWComboBoxUnits(5);
        }

        private void GWunitcomboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            GWComboBoxUnits(5);
        }

        private void NWunitcomboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            NWComboBoxUnits(6);
        }

        private void GWunitcomboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            GWComboBoxUnits(6);
        }

        private bool ValidatePrintingOptions(string selectedPrinter, int copies)
        {
            string errorMessage = "";

            // 檢查是否選擇了印表機
            if (string.IsNullOrEmpty(selectedPrinter))
            {
                errorMessage += "印表機尚未選擇\n"; // 如果沒有，添加相應的錯誤消息
            }

            // 檢查列印數量是否為0
            //if (copies == 0)
            //{
            //    errorMessage += "列印數量尚未選擇\n"; // 如果是，添加相應的錯誤消息
            //}

            // 如果存在錯誤消息
            if (!string.IsNullOrEmpty(errorMessage))
            {
                // 去除最後一個換行符，以便消息框更為整潔
                if (errorMessage.EndsWith("\n"))
                {
                    errorMessage = errorMessage.TrimEnd('\n');
                }

                // 顯示錯誤消息
                MessageBox.Show(errorMessage, "警告!");

                // 返回false，表示驗證未通過
                return false;
            }

            // 如果所有條件都符合，返回true
            return true;
        }

        private void PriviewPrintStart(int page, int padLeft)
        {
            currentPrintCount = 0; // 初始化當前列印計數為0
            double startQuantity = 0; // 初始化起始數量為0
                                      // 找到當前頁面的StarttextBox控件
            System.Windows.Forms.TextBox StarttextBox = (System.Windows.Forms.TextBox)tabControl.TabPages[page - 1].Controls["StarttextBox" + (page)];

            // 如果StarttextBox中有值，則解析並存儲在startQuantity中
            if (!string.IsNullOrEmpty(StarttextBox.Text))
            {
                startQuantity = double.Parse(StarttextBox.Text);
            }

            // 設置列印的標籤數量為1（因為這是預覽）
            btFormat.PrintSetup.IdenticalCopiesOfLabel = 1;
            string currentValue = ""; // 初始化當前數值字符串為空

            // 進行一次預覽列印
            for (double i = 1; i <= 1; i++)
            {
                currentValue = (startQuantity + currentPrintCount).ToString().PadLeft(padLeft, '0'); // 計算當前列印數值;
                btFormat.SubStrings["Current"].Value = currentValue; // 設置當前的列印數值

                // 顯示列印預覽對話框
                engine.Window.VisibleWindows = VisibleWindows.InteractiveDialogs;
                btFormat.PrintPreview.ShowDialog();

                // 更新當前列印數量
                currentPrintCount++;
            }
        }

        private void PrintStart(int page, int padLeft)
        {
            currentPrintCount = 0; // 初始化當前列印計數為0
            double startQuantity = 0; // 初始化起始數量為0
            double printQuantity = copies; // 從全局變量copies獲取列印數量

            // 找到相應頁面的StarttextBox控件並獲取其值
            System.Windows.Forms.TextBox StarttextBox = (System.Windows.Forms.TextBox)tabControl.TabPages[page - 1].Controls["StarttextBox" + (page)];
            if (!string.IsNullOrEmpty(StarttextBox.Text))
            {
                startQuantity = double.Parse(StarttextBox.Text);
            }

            // 設置列印的單一標籤數量為1
            btFormat.PrintSetup.IdenticalCopiesOfLabel = 1;

            // 找到相應頁面的EndtextBox控件
            System.Windows.Forms.TextBox EndtextBox = (System.Windows.Forms.TextBox)tabControl.TabPages[page - 1].Controls["EndtextBox" + (page)];

            // 新增 AddtextBox 的處理，用於獲取添加值
            System.Windows.Forms.TextBox AddtextBox = (System.Windows.Forms.TextBox)tabControl.TabPages[page - 1].Controls["AddtextBox" + (page)];
            double addValue = 0;
            if (!string.IsNullOrEmpty(AddtextBox.Text))
            {
                addValue = double.Parse(AddtextBox.Text);
            }

            string currentValue = ""; // 初始化當前列印數量的數值為空

            // 開始列印
            for (double i = 1; i <= printQuantity; i++)
            {
                for (double j = 0; j < SameCNOcopies; j++) // 同一CNO的副本數
                {
                    currentValue = (startQuantity + currentPrintCount).ToString().PadLeft(padLeft,'0'); // 計算當前列印數值
                    btFormat.SubStrings["Current"].Value = currentValue; // 設置列印數值

                    // 執行列印操作
                    btFormat.Print();
                }

                // 更新EndtextBox的數值，以反映當前列印數量
                EndtextBox.Text = (startQuantity + currentPrintCount).ToString();

                // 根據addValue更新當前列印數量
                currentPrintCount += addValue;
            }
        }

        private void DefaultSetting(int templateNumber)
        {
            selectedPrinter = Properties.Settings.Default.SelectedPrinter;
            SameCNOcopies = Properties.Settings.Default.Copies;
            tempPath = Path.Combine(Directory.GetCurrentDirectory(), "template", $"template{templateNumber}.btw");

        }

        private void Priviewbutton1_Click(object sender, EventArgs e)
        {
            DefaultSetting(1);

            if (!ValidatePrintingOptions(selectedPrinter, copies))
            {
                return;
            }

            // 初始化 Seagull BarTender 引擎
            engine.Start();

            // 開啟標籤文件
            btFormat = engine.Documents.Open(tempPath, selectedPrinter);
            // 參數說明：標籤路徑，印表機名稱

            // 設定標籤中的欄位值
            btFormat.SubStrings["Textbox"].Value = GraphictextBox1.Text;
            btFormat.SubStrings["NW"].Value = NWtextBox1.Text;
            if (NWunitcomboBox1.SelectedItem == null)
            {
                btFormat.SubStrings["NWunit"].Value = "";
            }
            else
            {
                btFormat.SubStrings["NWunit"].Value = " " + NWunitcomboBox1.SelectedItem.ToString();
            }
            btFormat.SubStrings["GW"].Value = GWtextBox1.Text;
            if (GWunitcomboBox1.SelectedItem == null)
            {
                btFormat.SubStrings["GWunit"].Value = "";
            }
            else
            {
                btFormat.SubStrings["GWunit"].Value = " " + GWunitcomboBox1.SelectedItem.ToString();
            }
            btFormat.SubStrings["SPEC"].Value = SPECtextBox1.Text;
            btFormat.SubStrings["Origin"].Value = OrigintextBox1.Text;
            btFormat.SubStrings["CNO"].Value = CNOtextBox1.Text;
            btFormat.SubStrings["Location"].Value = LocationtextBox1.Text;
            btFormat.SubStrings["LOTNumber"].Value = LOTtextBox1.Text;
            int padLeft = Convert.ToInt16(MD1_CBX_CNO.Text);
            PriviewPrintStart(1, padLeft);
        }

        private void Submitbutton1_Click(object sender, EventArgs e)
        {
            DefaultSetting(1);

            string errorMessage = "";
            if (string.IsNullOrEmpty(selectedPrinter))
            {
                errorMessage += "印表機尚未選擇\n";
            }
            //MessageBox.Show(selectedPrinter);
            if (copies == 0)
            {
                errorMessage += "列印數量尚未選擇\n";
            }
            //if (string.IsNullOrEmpty(GraphictextBox1.Text))
            //{
            //    errorMessage += "圖案內文字、 ";
            //}

            //if (string.IsNullOrEmpty(LocationtextBox1.Text))
            //{
            //    errorMessage += "Location、 ";
            //}

            //if (string.IsNullOrEmpty(CNOtextBox1.Text))
            //{
            //    errorMessage += "CNO、 ";
            //}

            //if (string.IsNullOrEmpty(OrigintextBox1.Text))
            //{
            //    errorMessage += "Origin、 ";
            //}

            //if (string.IsNullOrEmpty(SPECtextBox1.Text))
            //{
            //    errorMessage += "SPEC、 ";
            //}

            //if (string.IsNullOrEmpty(NWtextBox1.Text))
            //{
            //    errorMessage += "NW、 ";
            //}

            //if (string.IsNullOrEmpty(GWtextBox1.Text))
            //{
            //    errorMessage += "GW、 ";
            //}

            //if (NWunitcomboBox1.SelectedItem == null)
            //{
            //    errorMessage += "NWunit、 ";
            //}

            //if (GWunitcomboBox1.SelectedItem == null)
            //{
            //    errorMessage += "GWunit、 ";
            //}

            if (string.IsNullOrEmpty(LOTtextBox1.Text))
            {
                errorMessage += "LOTNumber、 ";
            }

            if (string.IsNullOrEmpty(StarttextBox1.Text))
            {
                errorMessage += "起始值、 ";
            }
            if (string.IsNullOrEmpty(AddtextBox1.Text))
            {
                errorMessage += "累加值、 ";
            }

            if (!string.IsNullOrEmpty(errorMessage))
            {
                if (errorMessage.EndsWith("、 "))
                {
                    errorMessage = errorMessage.TrimEnd(' ', '、'); // 刪除最後一個字元 "、"
                    errorMessage += "為空值";
                }
                MessageBox.Show(errorMessage, "警告!");
                return;
            }

            // 初始化 Seagull BarTender 引擎
            engine.Start();

            // 開啟標籤文件
            btFormat = engine.Documents.Open(tempPath, selectedPrinter);
            // 參數說明：標籤路徑，印表機名稱

            // 設定標籤中的欄位值
            btFormat.SubStrings["Textbox"].Value = GraphictextBox1.Text;
            btFormat.SubStrings["NW"].Value = NWtextBox1.Text;
            btFormat.SubStrings["NWunit"].Value = " " + NWunitcomboBox1.Text;
            btFormat.SubStrings["GW"].Value = GWtextBox1.Text;
            btFormat.SubStrings["GWunit"].Value = " " + GWunitcomboBox1.Text;
            btFormat.SubStrings["SPEC"].Value = SPECtextBox1.Text;
            btFormat.SubStrings["Origin"].Value = OrigintextBox1.Text;
            btFormat.SubStrings["CNO"].Value = CNOtextBox1.Text;
            btFormat.SubStrings["Location"].Value = LocationtextBox1.Text;
            btFormat.SubStrings["LOTNumber"].Value = LOTtextBox1.Text;
            int padLeft = Convert.ToInt16(MD1_CBX_CNO.Text);
            
            PrintStart(1, padLeft);

            engine.Stop();

            Properties.Settings.Default.LOTNum1 = LOTtextBox1.Text;
            Properties.Settings.Default.Save();
        }

        private void Resetbutton1_Click(object sender, EventArgs e)
        {
            engine.Stop();
            // 清除所有的 TextBox
            GraphictextBox1.Clear();
            LocationtextBox1.Clear();
            CNOtextBox1.Clear();
            OrigintextBox1.Clear();
            SPECtextBox1.Clear();
            NWtextBox1.Clear();
            GWtextBox1.Clear();
            LOTtextBox1.Clear();
            StarttextBox1.Clear();
            EndtextBox1.Clear();
            GWunitcomboBox1.SelectedIndex = -1;
            NWunitcomboBox1.SelectedIndex = -1;
            PrintQuantitynumericUpDown1.Value = 0;
            AddtextBox1.Text = "1";

            if (string.IsNullOrEmpty(OrigintextBox1.Text))
            {
                // 檢查OrigintextBox是否為空，若是則恢復預設文字和外觀
                OrigintextBox1.Text = "MADE IN";
                OrigintextBox1.ForeColor = SystemColors.GrayText;
                OrigintextBox1.TextAlign = HorizontalAlignment.Left;
            }
        }

        private void Priviewbutton2_Click(object sender, EventArgs e)
        {
            DefaultSetting(2);

            if (!ValidatePrintingOptions(selectedPrinter, copies))
            {
                return;
            }

            // 初始化 Seagull BarTender 引擎
            engine.Start();

            // 開啟標籤文件
            btFormat = engine.Documents.Open(tempPath, selectedPrinter);
            // 參數說明：標籤路徑，印表機名稱

            // 設定標籤中的欄位值
            btFormat.SubStrings["LOTNumber"].Value = LOTtextBox2.Text;
            btFormat.SubStrings["Grade"].Value = GradetextBox2.Text;
            btFormat.SubStrings["SPEC"].Value = SPECtextBox2.Text;
            btFormat.SubStrings["Quantity"].Value = QuantitytextBox2.Text;
            btFormat.SubStrings["NW"].Value = NWtextBox2.Text;
            if (NWunitcomboBox2.SelectedItem == null)
            {
                btFormat.SubStrings["NWunit"].Value = "";
            }
            else
            {
                btFormat.SubStrings["NWunit"].Value = NWunitcomboBox2.SelectedItem.ToString();
            }
            btFormat.SubStrings["GW"].Value = GWtextBox2.Text;
            if (GWunitcomboBox2.SelectedItem == null)
            {
                btFormat.SubStrings["GWunit"].Value = "";
            }
            else
            {
                btFormat.SubStrings["GWunit"].Value = GWunitcomboBox2.SelectedItem.ToString();
            }
            btFormat.SubStrings["CNO"].Value = CNOtextBox2.Text;
            int padLeft = Convert.ToInt16(MD2_CBX_CNO.Text);

            PriviewPrintStart(2, padLeft);
        }

        private void Submitbutton2_Click(object sender, EventArgs e)
        {
            DefaultSetting(2);

            string errorMessage = "";
            if (string.IsNullOrEmpty(selectedPrinter))
            {
                errorMessage += "印表機尚未選擇\n";
            }
            //MessageBox.Show(selectedPrinter);
            if (copies == 0)
            {
                errorMessage += "列印數量尚未選擇\n";
            }
            //確保輸入不為空值
            //if (string.IsNullOrEmpty(LOTtextBox2.Text))
            //{
            //    errorMessage += "LOTNumber、 ";
            //}

            //if (string.IsNullOrEmpty(GradetextBox2.Text))
            //{
            //    errorMessage += "Grade、 ";
            //}

            //if (string.IsNullOrEmpty(SPECtextBox2.Text))
            //{
            //    errorMessage += "SPEC、 ";
            //}

            //if (string.IsNullOrEmpty(QuantitytextBox2.Text))
            //{
            //    errorMessage += "Quantity、 ";
            //}

            //if (string.IsNullOrEmpty(NWtextBox2.Text))
            //{
            //    errorMessage += "NW、 ";
            //}

            //if (string.IsNullOrEmpty(GWtextBox2.Text))
            //{
            //    errorMessage += "GW、 ";
            //}

            //if (NWunitcomboBox2.SelectedItem == null)
            //{
            //    errorMessage += "NWunit、 ";
            //}

            //if (GWunitcomboBox2.SelectedItem == null)
            //{
            //    errorMessage += "GWunit、 ";
            //}

            //if (string.IsNullOrEmpty(CNOtextBox2.Text))
            //{
            //    errorMessage += "CNO、 ";
            //}

            if (string.IsNullOrEmpty(StarttextBox2.Text))
            {
                errorMessage += "起始值、 ";
            }
            if (string.IsNullOrEmpty(AddtextBox2.Text))
            {
                errorMessage += "累加值、 ";
            }

            if (!string.IsNullOrEmpty(errorMessage))
            {
                if (errorMessage.EndsWith("、 "))
                {
                    errorMessage = errorMessage.TrimEnd(' ', '、'); // 刪除最後一個字元 "、"
                    errorMessage += "為空值";
                }
                MessageBox.Show(errorMessage, "警告!");
                return;
            }

            // 初始化 Seagull BarTender 引擎
            engine.Start();

            // 開啟標籤文件
            btFormat = engine.Documents.Open(tempPath, selectedPrinter);
            // 參數說明：標籤路徑，印表機名稱

            // 設定標籤中的欄位值
            btFormat.SubStrings["LOTNumber"].Value = LOTtextBox2.Text;
            btFormat.SubStrings["Grade"].Value = GradetextBox2.Text;
            btFormat.SubStrings["SPEC"].Value = SPECtextBox2.Text;
            btFormat.SubStrings["Quantity"].Value = QuantitytextBox2.Text;
            btFormat.SubStrings["NW"].Value = NWtextBox2.Text;
            btFormat.SubStrings["NWunit"].Value = NWunitcomboBox2.Text;
            btFormat.SubStrings["GW"].Value = GWtextBox2.Text;
            btFormat.SubStrings["GWunit"].Value = GWunitcomboBox2.Text;
            btFormat.SubStrings["CNO"].Value = CNOtextBox2.Text;
            int padLeft = Convert.ToInt16(MD2_CBX_CNO.Text);
            PrintStart(2, padLeft);

            engine.Stop();

            Properties.Settings.Default.LOTNum2 = LOTtextBox2.Text;
            Properties.Settings.Default.Save();
        }

        private void Resetbutton2_Click(object sender, EventArgs e)
        {
            engine.Stop();
            // 清除所有的 TextBox
            CNOtextBox2.Clear();
            GradetextBox2.Clear();
            SPECtextBox2.Clear();
            QuantitytextBox2.Clear();
            NWtextBox2.Clear();
            GWtextBox2.Clear();
            LOTtextBox2.Clear();
            StarttextBox2.Clear();
            EndtextBox2.Clear();
            GWunitcomboBox2.SelectedIndex = -1;
            NWunitcomboBox2.SelectedIndex = -1;
            PrintQuantitynumericUpDown2.Value = 0;
            AddtextBox2.Text = "1";
        }

        private void Priviewbutton3_Click(object sender, EventArgs e)
        {
            DefaultSetting(3);

            if (!ValidatePrintingOptions(selectedPrinter, copies))
            {
                return;
            }

            // 初始化 Seagull BarTender 引擎
            engine.Start();

            // 開啟標籤文件
            btFormat = engine.Documents.Open(tempPath, selectedPrinter);
            // 參數說明：標籤路徑，印表機名稱

            // 設定標籤中的欄位值
            btFormat.SubStrings["Textbox"].Value = GraphictextBox3.Text;
            btFormat.SubStrings["MATERIAL"].Value = MaterialtextBox3.Text;
            btFormat.SubStrings["CNO"].Value = CNOtextBox3.Text;
            //btFormat.SubStrings["SPEC"].Value = SPECtextBox3.Text;
            btFormat.SubStrings["NW"].Value = NWtextBox3.Text;
            if (NWunitcomboBox3.SelectedItem == null)
            {
                btFormat.SubStrings["NWunit"].Value = "";
            }
            else
            {
                btFormat.SubStrings["NWunit"].Value = " " + NWunitcomboBox3.SelectedItem.ToString();
            }
            btFormat.SubStrings["GW"].Value = GWtextBox3.Text;
            if (GWunitcomboBox3.SelectedItem == null)
            {
                btFormat.SubStrings["GWunit"].Value = "";
            }
            else
            {
                btFormat.SubStrings["GWunit"].Value = " " + GWunitcomboBox3.SelectedItem.ToString();
            }
            btFormat.SubStrings["Origin"].Value = OrigintextBox3.Text;
            int padLeft = Convert.ToInt16(MD3_CBX_CNO.Text);
            PriviewPrintStart(3, padLeft);
        }

        private void Submitbutton3_Click(object sender, EventArgs e)
        {
            DefaultSetting(3);

            string errorMessage = "";
            if (string.IsNullOrEmpty(selectedPrinter))
            {
                errorMessage += "印表機尚未選擇\n";
            }
            if (copies == 0)
            {
                errorMessage += "列印數量尚未選擇\n";
            }
            //確保輸入不為空值
            //if (string.IsNullOrEmpty(GraphictextBox3.Text))
            //{
            //    errorMessage += "圖案內文字、 ";
            //}
            //if (string.IsNullOrEmpty(MaterialtextBox3.Text))
            //{
            //    errorMessage += "Material、 ";
            //}
            if (string.IsNullOrEmpty(StarttextBox3.Text))
            {
                errorMessage += "起始值、 ";
            }
            if (string.IsNullOrEmpty(AddtextBox3.Text))
            {
                errorMessage += "累加值、 ";
            }
            //if (string.IsNullOrEmpty(NWtextBox3.Text))
            //{
            //    errorMessage += "NW、 ";
            //}

            //if (string.IsNullOrEmpty(GWtextBox3.Text))
            //{
            //    errorMessage += "GW、 ";
            //}

            //if (NWunitcomboBox3.SelectedItem == null)
            //{
            //    errorMessage += "NWunit、 ";
            //}

            //if (GWunitcomboBox3.SelectedItem == null)
            //{
            //    errorMessage += "GWunit、 ";
            //}
            //if (string.IsNullOrEmpty(OrigintextBox3.Text))
            //{
            //    errorMessage += "Origin、 ";
            //}
            if (!string.IsNullOrEmpty(errorMessage))
            {
                if (errorMessage.EndsWith("、 "))
                {
                    errorMessage = errorMessage.TrimEnd(' ', '、'); // 刪除最後一個字元 "、"
                    errorMessage += "為空值";
                }
                MessageBox.Show(errorMessage, "警告!");
                return;
            }

            // 初始化 Seagull BarTender 引擎
            engine.Start();

            // 開啟標籤文件
            btFormat = engine.Documents.Open(tempPath, selectedPrinter);
            // 參數說明：標籤路徑，印表機名稱

            // 設定標籤中的欄位值
            btFormat.SubStrings["Textbox"].Value = GraphictextBox3.Text;
            btFormat.SubStrings["MATERIAL"].Value = MaterialtextBox3.Text;
            btFormat.SubStrings["CNO"].Value = CNOtextBox3.Text;
            //btFormat.SubStrings["SPEC"].Value = SPECtextBox3.Text;
            btFormat.SubStrings["NW"].Value = NWtextBox3.Text;
            btFormat.SubStrings["NWunit"].Value = " " + NWunitcomboBox3.Text;
            btFormat.SubStrings["GW"].Value = GWtextBox3.Text;
            btFormat.SubStrings["GWunit"].Value = " " + GWunitcomboBox3.Text;
            btFormat.SubStrings["Origin"].Value = OrigintextBox3.Text;
            int padLeft = Convert.ToInt16(MD3_CBX_CNO.Text);
            PrintStart(3, padLeft);

            engine.Stop();

        }

        private void Resetbutton3_Click(object sender, EventArgs e)
        {
            engine.Stop();
            // 清除所有的 TextBox
            GraphictextBox3.Clear();
            MaterialtextBox3.Clear();
            StarttextBox3.Clear();
            EndtextBox3.Clear();
            //SPECtextBox3.Clear();
            OrigintextBox3.Clear();
            NWtextBox3.Clear();
            GWtextBox3.Clear();
            GWunitcomboBox3.SelectedIndex = -1;
            NWunitcomboBox3.SelectedIndex = -1;
            PrintQuantitynumericUpDown3.Value = 0;
            AddtextBox3.Text = "1";

            if (string.IsNullOrEmpty(OrigintextBox3.Text))
            {
                // 檢查OrigintextBox是否為空，若是則恢復預設文字和外觀
                OrigintextBox3.Text = "MADE IN";
                OrigintextBox3.ForeColor = SystemColors.GrayText;
                OrigintextBox3.TextAlign = HorizontalAlignment.Left;
            }
        }

        private void Priviewbutton4_Click(object sender, EventArgs e)
        {
            DefaultSetting(4);

            if (!ValidatePrintingOptions(selectedPrinter, copies))
            {
                return;
            }

            // 初始化 Seagull BarTender 引擎
            engine.Start();

            // 開啟標籤文件
            btFormat = engine.Documents.Open(tempPath, selectedPrinter);
            // 參數說明：標籤路徑，印表機名稱

            // 設定標籤中的欄位值
            btFormat.SubStrings["SPEC"].Value = SPECtextBox4.Text;
            btFormat.SubStrings["CNO"].Value = CNOtextBox4.Text;
            btFormat.SubStrings["ORIGIN"].Value = OrigintextBox4.Text;
            btFormat.SubStrings["NW"].Value = NWtextBox4.Text;
            if (NWunitcomboBox4.SelectedItem == null)
            {
                btFormat.SubStrings["NWunit"].Value = "";
            }
            else
            {
                btFormat.SubStrings["NWunit"].Value = "  " + NWunitcomboBox4.SelectedItem.ToString();
            }
            btFormat.SubStrings["GW"].Value = GWtextBox4.Text;
            if (GWunitcomboBox4.SelectedItem == null)
            {
                btFormat.SubStrings["GWunit"].Value = "";
            }
            else
            {
                btFormat.SubStrings["GWunit"].Value = "  " + GWunitcomboBox4.SelectedItem.ToString();
            }
            btFormat.SubStrings["LOTNumber"].Value = LOTtextBox4.Text;
            int padLeft = Convert.ToInt16(MD4_CBX_CNO.Text);
            PriviewPrintStart(4, padLeft);
        }

        private void Submitbutton4_Click(object sender, EventArgs e)
        {
            DefaultSetting(4);

            string errorMessage = "";
            if (string.IsNullOrEmpty(selectedPrinter))
            {
                errorMessage += "印表機尚未選擇\n";
            }
            //MessageBox.Show(selectedPrinter);
            if (copies == 0)
            {
                errorMessage += "列印數量尚未選擇\n";
            }
            //確保輸入不為空值
            //if (string.IsNullOrEmpty(SPECtextBox4.Text))
            //{
            //    errorMessage += "SPEC、 ";
            //}
            //if (string.IsNullOrEmpty(CNOtextBox4.Text))
            //{
            //    errorMessage += "CNO、 ";
            //}

            //if (string.IsNullOrEmpty(OrigintextBox4.Text))
            //{
            //    errorMessage += "Origin、 ";
            //}
            if (string.IsNullOrEmpty(StarttextBox4.Text))
            {
                errorMessage += "起始值、 ";
            }
            if (string.IsNullOrEmpty(AddtextBox4.Text))
            {
                errorMessage += "累加值、 ";
            }
            //if (string.IsNullOrEmpty(NWtextBox4.Text))
            //{
            //    errorMessage += "NW、 ";
            //}

            //if (string.IsNullOrEmpty(GWtextBox4.Text))
            //{
            //    errorMessage += "GW、 ";
            //}

            //if (NWunitcomboBox4.SelectedItem == null)
            //{
            //    errorMessage += "NWunit、 ";
            //}

            //if (GWunitcomboBox4.SelectedItem == null)
            //{
            //    errorMessage += "GWunit、 ";
            //}
            //if (string.IsNullOrEmpty(LOTtextBox4.Text))
            //{
            //    errorMessage += "LOTNumber、 ";
            //}
            if (!string.IsNullOrEmpty(errorMessage))
            {
                if (errorMessage.EndsWith("、 "))
                {
                    errorMessage = errorMessage.TrimEnd(' ', '、'); // 刪除最後一個字元 "、"
                    errorMessage += "為空值";
                }
                MessageBox.Show(errorMessage, "警告!");
                return;
            }

            // 初始化 Seagull BarTender 引擎
            engine.Start();

            // 開啟標籤文件
            btFormat = engine.Documents.Open(tempPath, selectedPrinter);
            // 參數說明：標籤路徑，印表機名稱

            // 設定標籤中的欄位值
            btFormat.SubStrings["SPEC"].Value = SPECtextBox4.Text;
            btFormat.SubStrings["CNO"].Value = CNOtextBox4.Text;
            btFormat.SubStrings["ORIGIN"].Value = OrigintextBox4.Text;
            btFormat.SubStrings["NW"].Value = NWtextBox4.Text;
            btFormat.SubStrings["NWunit"].Value = "  " + NWunitcomboBox4.Text;
            btFormat.SubStrings["GW"].Value = GWtextBox4.Text;
            btFormat.SubStrings["GWunit"].Value = "  " + GWunitcomboBox4.Text;
            btFormat.SubStrings["LOTNumber"].Value = LOTtextBox4.Text;
            int padLeft = Convert.ToInt16(MD4_CBX_CNO.Text);
            PrintStart(4, padLeft);

            engine.Stop();

            Properties.Settings.Default.LOTNum4 = LOTtextBox4.Text;
            Properties.Settings.Default.Save();
        }

        private void Resetbutton4_Click(object sender, EventArgs e)
        {
            engine.Stop();
            // 清除所有的 TextBox
            CNOtextBox4.Clear();
            OrigintextBox4.Clear();
            SPECtextBox4.Clear();
            NWtextBox4.Clear();
            GWtextBox4.Clear();
            LOTtextBox4.Clear();
            StarttextBox4.Clear();
            EndtextBox4.Clear();
            GWunitcomboBox4.SelectedIndex = -1;
            NWunitcomboBox4.SelectedIndex = -1;
            PrintQuantitynumericUpDown4.Value = 0;
            AddtextBox4.Text = "1";

            if (string.IsNullOrEmpty(OrigintextBox4.Text))
            {
                // 檢查OrigintextBox是否為空，若是則恢復預設文字和外觀
                OrigintextBox4.Text = "MADE IN";
                OrigintextBox4.ForeColor = SystemColors.GrayText;
                OrigintextBox4.TextAlign = HorizontalAlignment.Left;
            }
        }

        private void Priviewbutton5_Click(object sender, EventArgs e)
        {
            DefaultSetting(5);

            if (!ValidatePrintingOptions(selectedPrinter, copies))
            {
                return;
            }

            // 初始化 Seagull BarTender 引擎
            engine.Start();

            // 開啟標籤文件
            btFormat = engine.Documents.Open(tempPath, selectedPrinter);
            // 參數說明：標籤路徑，印表機名稱

            // 設定標籤中的欄位值
            btFormat.SubStrings["LOCATION1"].Value = LocationtextBox15.Text;
            btFormat.SubStrings["LOCATION2"].Value = LocationtextBox25.Text;
            btFormat.SubStrings["CNO"].Value = CNOtextBox5.Text;
            btFormat.SubStrings["ORIGIN"].Value = OrigintextBox5.Text;
            btFormat.SubStrings["SPEC"].Value = SPECtextBox5.Text;
            if (NWunitcomboBox5.SelectedItem == null)
            {
                btFormat.SubStrings["NWunit"].Value = "";
            }
            else
            {
                btFormat.SubStrings["NWunit"].Value = " " + NWunitcomboBox5.SelectedItem.ToString();
            }
            btFormat.SubStrings["GW"].Value = GWtextBox5.Text;
            if (GWunitcomboBox5.SelectedItem == null)
            {
                btFormat.SubStrings["GWunit"].Value = "";
            }
            else
            {
                btFormat.SubStrings["GWunit"].Value = " " + GWunitcomboBox5.SelectedItem.ToString();
            }
            btFormat.SubStrings["LOTNumber"].Value = LOTtextBox5.Text;
            int padLeft = Convert.ToInt16(MD5_CBX_CNO.Text);
            PriviewPrintStart(5, padLeft);
        }

        private void Submitbutton5_Click(object sender, EventArgs e)
        {
            DefaultSetting(5);

            string errorMessage = "";
            if (string.IsNullOrEmpty(selectedPrinter))
            {
                errorMessage += "印表機尚未選擇\n";
            }
            if (copies == 0)
            {
                errorMessage += "列印數量尚未選擇\n";
            }
            //if (string.IsNullOrEmpty(LocationtextBox15.Text))
            //{
            //    errorMessage += "Location1、 ";
            //}
            //if (string.IsNullOrEmpty(LocationtextBox25.Text))
            //{
            //    errorMessage += "Location2、 ";
            //}
            //if (string.IsNullOrEmpty(CNOtextBox5.Text))
            //{
            //    errorMessage += "CNO、 ";
            //}

            //if (string.IsNullOrEmpty(OrigintextBox5.Text))
            //{
            //    errorMessage += "Origin、 ";
            //}
            if (string.IsNullOrEmpty(StarttextBox5.Text))
            {
                errorMessage += "起始值、 ";
            }
            if (string.IsNullOrEmpty(AddtextBox5.Text))
            {
                errorMessage += "累加值、 ";
            }
            //if (string.IsNullOrEmpty(SPECtextBox5.Text))
            //{
            //    errorMessage += "SPEC、 ";
            //}

            //if (string.IsNullOrEmpty(NWtextBox5.Text))
            //{
            //    errorMessage += "NW、 ";
            //}

            //if (string.IsNullOrEmpty(GWtextBox5.Text))
            //{
            //    errorMessage += "GW、 ";
            //}

            //if (NWunitcomboBox5.SelectedItem == null)
            //{
            //    errorMessage += "NWunit、 ";
            //}

            //if (GWunitcomboBox5.SelectedItem == null)
            //{
            //    errorMessage += "GWunit、 ";
            //}
            //if (string.IsNullOrEmpty(LOTtextBox5.Text))
            //{
            //    errorMessage += "LOTNumber、 ";
            //}
            if (!string.IsNullOrEmpty(errorMessage))
            {
                if (errorMessage.EndsWith("、 "))
                {
                    errorMessage = errorMessage.TrimEnd(' ', '、'); // 刪除最後一個字元 "、"
                    errorMessage += "為空值";
                }
                MessageBox.Show(errorMessage, "警告!");
                return;
            }

            // 初始化 Seagull BarTender 引擎

            engine.Start();

            // 開啟標籤文件
            btFormat = engine.Documents.Open(tempPath, selectedPrinter);
            // 參數說明：標籤路徑，印表機名稱

            // 設定標籤中的欄位值
            btFormat.SubStrings["LOCATION1"].Value = LocationtextBox15.Text;
            btFormat.SubStrings["LOCATION2"].Value = LocationtextBox25.Text;
            btFormat.SubStrings["CNO"].Value = CNOtextBox5.Text;
            btFormat.SubStrings["ORIGIN"].Value = OrigintextBox5.Text;
            btFormat.SubStrings["SPEC"].Value = SPECtextBox5.Text;
            btFormat.SubStrings["NW"].Value = NWtextBox5.Text;
            btFormat.SubStrings["NWunit"].Value = " " + NWunitcomboBox5.Text;
            btFormat.SubStrings["GW"].Value = GWtextBox5.Text;
            btFormat.SubStrings["GWunit"].Value = " " + GWunitcomboBox5.Text;
            btFormat.SubStrings["LOTNumber"].Value = LOTtextBox5.Text;
            int padLeft = Convert.ToInt16(MD5_CBX_CNO.Text);
            PrintStart(5, padLeft);

            engine.Stop();

            Properties.Settings.Default.LOTNum5 = LOTtextBox5.Text;
            Properties.Settings.Default.Save();
        }

        private void Resetbutton5_Click(object sender, EventArgs e)
        {
            engine.Stop();
            // 清除所有的 TextBox
            LocationtextBox15.Clear();
            LocationtextBox25.Clear();
            CNOtextBox5.Clear();
            OrigintextBox5.Clear();
            SPECtextBox5.Clear();
            NWtextBox5.Clear();
            GWtextBox5.Clear();
            LOTtextBox5.Clear();
            StarttextBox5.Clear();
            EndtextBox5.Clear();
            GWunitcomboBox5.SelectedIndex = -1;
            NWunitcomboBox5.SelectedIndex = -1;
            PrintQuantitynumericUpDown5.Value = 0;
            AddtextBox5.Text = "1";

            if (string.IsNullOrEmpty(OrigintextBox5.Text))
            {
                // 檢查OrigintextBox是否為空，若是則恢復預設文字和外觀
                OrigintextBox5.Text = "MADE IN";
                OrigintextBox5.ForeColor = SystemColors.GrayText;
                OrigintextBox5.TextAlign = HorizontalAlignment.Left;
            }
        }

        private void Priviewbutton6_Click(object sender, EventArgs e)
        {
            DefaultSetting(6);

            if (!ValidatePrintingOptions(selectedPrinter, copies))
            {
                return;
            }

            // 初始化 Seagull BarTender 引擎
            engine.Start();

            // 開啟標籤文件
            btFormat = engine.Documents.Open(tempPath, selectedPrinter);
            // 參數說明：標籤路徑，印表機名稱

            // 設定標籤中的欄位值
            btFormat.SubStrings["LOTNumber"].Value = LOTtextBox6.Text;
            btFormat.SubStrings["YarnCount"].Value = YarnCounttextBox6.Text;
            btFormat.SubStrings["LinearDensity"].Value = LinearDensitynumericUpDown6.Value.ToString();
            btFormat.SubStrings["LinearDensityUnit"].Value = LinearDensityUnitcomboBox6.SelectedItem.ToString();
            btFormat.SubStrings["Importer"].Value = ImportertextBox6.Text;
            btFormat.SubStrings["CNPJ"].Value = CNPJtextBox6.Text;
            btFormat.SubStrings["Manufacturer"].Value = ManufacturertextBox6.Text;
            btFormat.SubStrings["ORIGIN"].Value = OrigintextBox6.Text;
            btFormat.SubStrings["MATERIAL"].Value = MaterialtextBox6.Text;
            btFormat.SubStrings["CNO"].Value = CNOtextBox6.Text;
            if (NWunitcomboBox6.SelectedItem == null)
            {
                btFormat.SubStrings["NWunit"].Value = "";
            }
            else
            {
                btFormat.SubStrings["NWunit"].Value = "  " + NWunitcomboBox6.SelectedItem.ToString();
            }
            btFormat.SubStrings["GW"].Value = GWtextBox6.Text;
            if (GWunitcomboBox6.SelectedItem == null)
            {
                btFormat.SubStrings["GWunit"].Value = "";
            }
            else
            {
                btFormat.SubStrings["GWunit"].Value = "  " + GWunitcomboBox6.SelectedItem.ToString();
            }
            btFormat.SubStrings["LOTNumber"].Value = LOTtextBox6.Text;
            int padLeft = Convert.ToInt16(MD6_CBX_CNO.Text);
            PriviewPrintStart(6, padLeft);
        }

        private void Submitbutton6_Click(object sender, EventArgs e)
        {
            DefaultSetting(6);

            string errorMessage = "";
            if (string.IsNullOrEmpty(selectedPrinter))
            {
                errorMessage += "印表機尚未選擇\n";
            }
            if (copies == 0)
            {
                errorMessage += "列印數量尚未選擇\n";
            }
            //確保輸入不為空值
            //if (string.IsNullOrEmpty(LOTtextBox6.Text))
            //{
            //    errorMessage += "LOTNumber、 ";
            //}
            //if (string.IsNullOrEmpty(YarnCounttextBox6.Text))
            //{
            //    errorMessage += "YarnCount、 ";
            //}
            if (LinearDensitynumericUpDown6.Value == 0)
            {
                errorMessage += "LinearDensity、 ";
            }
            if (LinearDensityUnitcomboBox6.SelectedItem == null)
            {
                errorMessage += "LinearDensityUnit、 ";
            }
            //if (string.IsNullOrEmpty(ImportertextBox6.Text))
            //{
            //    errorMessage += "Importer、 ";
            //}
            //if (string.IsNullOrEmpty(CNPJtextBox6.Text))
            //{
            //    errorMessage += "CNPJ、 ";
            //}
            //if (string.IsNullOrEmpty(ManufacturertextBox6.Text))
            //{
            //    errorMessage += "Manufacturer、 ";
            //}
            //if (string.IsNullOrEmpty(OrigintextBox6.Text))
            //{
            //    errorMessage += "Origin、 ";
            //}
            //if (string.IsNullOrEmpty(MaterialtextBox6.Text))
            //{
            //    errorMessage += "Material、 ";
            //}
            //if (string.IsNullOrEmpty(LOTtextBox6.Text))
            //{
            //    errorMessage += "LOTNumber、 ";
            //}
            //if (string.IsNullOrEmpty(CNOtextBox6.Text))
            //{
            //    errorMessage += "CNO、 ";
            //}
            if (string.IsNullOrEmpty(StarttextBox6.Text))
            {
                errorMessage += "起始值、 ";
            }
            if (string.IsNullOrEmpty(AddtextBox6.Text))
            {
                errorMessage += "累加值、 ";
            }
            //if (string.IsNullOrEmpty(NWtextBox6.Text))
            //{
            //    errorMessage += "NW、 ";
            //}

            //if (string.IsNullOrEmpty(GWtextBox6.Text))
            //{
            //    errorMessage += "GW、 ";
            //}

            //if (NWunitcomboBox6.SelectedItem == null)
            //{
            //    errorMessage += "NWunit、 ";
            //}

            //if (GWunitcomboBox6.SelectedItem == null)
            //{
            //    errorMessage += "GWunit、 ";
            //}
            if (!string.IsNullOrEmpty(errorMessage))
            {
                if (errorMessage.EndsWith("、 "))
                {
                    errorMessage = errorMessage.TrimEnd(' ', '、'); // 刪除最後一個字元 "、"
                    errorMessage += "為空值";
                }
                MessageBox.Show(errorMessage, "警告!");
                return;
            }

            // 初始化 Seagull BarTender 引擎
            engine.Start();

            // 開啟標籤文件
            btFormat = engine.Documents.Open(tempPath, selectedPrinter);
            // 參數說明：標籤路徑，印表機名稱

            // 設定標籤中的欄位值
            btFormat.SubStrings["LOTNumber"].Value = LOTtextBox6.Text;
            btFormat.SubStrings["YarnCount"].Value = YarnCounttextBox6.Text;
            btFormat.SubStrings["LinearDensity"].Value = LinearDensitynumericUpDown6.Value.ToString();
            btFormat.SubStrings["LinearDensityUnit"].Value = LinearDensityUnitcomboBox6.SelectedItem.ToString();
            btFormat.SubStrings["Importer"].Value = ImportertextBox6.Text;
            btFormat.SubStrings["CNPJ"].Value = CNPJtextBox6.Text;
            btFormat.SubStrings["Manufacturer"].Value = ManufacturertextBox6.Text;
            btFormat.SubStrings["ORIGIN"].Value = OrigintextBox6.Text;
            btFormat.SubStrings["MATERIAL"].Value = MaterialtextBox6.Text;
            btFormat.SubStrings["CNO"].Value = CNOtextBox6.Text;
            btFormat.SubStrings["NW"].Value = NWtextBox6.Text;
            btFormat.SubStrings["NWunit"].Value = "  " + NWunitcomboBox6.Text;
            btFormat.SubStrings["GW"].Value = GWtextBox6.Text;
            btFormat.SubStrings["GWunit"].Value = "  " + GWunitcomboBox6.Text;
            int padLeft = Convert.ToInt16(MD6_CBX_CNO.Text);
            PrintStart(6, padLeft);

            engine.Stop();

            Properties.Settings.Default.LOTNum6 = LOTtextBox6.Text;
            Properties.Settings.Default.Save();
        }

        private void Resetbutton6_Click(object sender, EventArgs e)
        {
            engine.Stop();
            // 清除所有的 TextBox
            LOTtextBox6.Clear();
            YarnCounttextBox6.Clear();
            ImportertextBox6.Clear();
            ManufacturertextBox6.Clear();
            OrigintextBox6.Clear();
            CNPJtextBox6.Clear();
            StarttextBox6.Clear();
            EndtextBox6.Clear();
            MaterialtextBox6.Clear();
            LinearDensityUnitcomboBox6.SelectedItem = -1;
            GWunitcomboBox6.SelectedIndex = -1;
            NWunitcomboBox6.SelectedIndex = -1;
            LinearDensityUnitcomboBox6.SelectedIndex = -1;
            LinearDensitynumericUpDown6.Value = LinearDensitynumericUpDown6.Minimum;
            NWtextBox6.Clear();
            GWtextBox6.Clear();
            PrintQuantitynumericUpDown6.Value = 0;
            AddtextBox6.Text = "1";

            if (string.IsNullOrEmpty(OrigintextBox6.Text))
            {
                // 檢查OrigintextBox是否為空，若是則恢復預設文字和外觀
                OrigintextBox6.Text = "MADE IN";
                OrigintextBox6.ForeColor = SystemColors.GrayText;
                OrigintextBox6.TextAlign = HorizontalAlignment.Left;
            }
        }

        private void Priviewbutton7_Click(object sender, EventArgs e)
        {
            DefaultSetting(7);

            if (!ValidatePrintingOptions(selectedPrinter, copies))
            {
                return;
            }

            // 初始化 Seagull BarTender 引擎
            engine.Start();

            // 開啟標籤文件
            btFormat = engine.Documents.Open(tempPath, selectedPrinter);
            // 參數說明：標籤路徑，印表機名稱

            // 設定標籤中的欄位值
            btFormat.SubStrings["Importer"].Value = ImportertextBox7.Text;
            btFormat.SubStrings["SerialNumber"].Value = SNtextBox7.Text;
            btFormat.SubStrings["PurchaseOrder"].Value = POtextBox7.Text;
            btFormat.SubStrings["MATERIAL"].Value = MaterialtextBox7.Text;
            btFormat.SubStrings["LOTNumber"].Value = LOTtextBox7.Text;
            btFormat.SubStrings["CNO"].Value = CartontextBox7.Text;
            btFormat.SubStrings["ORIGIN"].Value = OrigintextBox7.Text;
            int padLeft = Convert.ToInt16(MD7_CBX_CNO.Text);
            PriviewPrintStart(7, padLeft);
        }

        private void Submitbutton7_Click(object sender, EventArgs e)
        {
            DefaultSetting(7);

            string errorMessage = "";
            if (string.IsNullOrEmpty(selectedPrinter))
            {
                errorMessage += "印表機尚未選擇\n";
            }
            if (copies == 0)
            {
                errorMessage += "列印數量尚未選擇\n";
            }
            //if (string.IsNullOrEmpty(ImportertextBox7.Text))
            //{
            //    errorMessage += "Importer、 ";
            //}
            //if (string.IsNullOrEmpty(SNtextBox7.Text))
            //{
            //    errorMessage += "SerialNumber、 ";
            //}
            //if (string.IsNullOrEmpty(POtextBox7.Text))
            //{
            //    errorMessage += "PurchaseOrder、 ";
            //}
            //if (string.IsNullOrEmpty(MaterialtextBox7.Text))
            //{
            //    errorMessage += "Material、 ";
            //}
            //if (string.IsNullOrEmpty(LOTtextBox7.Text))
            //{
            //    errorMessage += "LOTNumber、 ";
            //}
            //if (string.IsNullOrEmpty(CartontextBox7.Text))
            //{
            //    errorMessage += "Carton、 ";
            //}
            if (string.IsNullOrEmpty(StarttextBox7.Text))
            {
                errorMessage += "起始值、 ";
            }
            if (string.IsNullOrEmpty(AddtextBox7.Text))
            {
                errorMessage += "累加值、 ";
            }
            //if (string.IsNullOrEmpty(OrigintextBox7.Text))
            //{
            //    errorMessage += "Origin、 ";
            //}
            if (!string.IsNullOrEmpty(errorMessage))
            {
                if (errorMessage.EndsWith("、 "))
                {
                    errorMessage = errorMessage.TrimEnd(' ', '、'); // 刪除最後一個字元 "、"
                    errorMessage += "為空值";
                }
                MessageBox.Show(errorMessage, "警告!");
                return;
            }

            // 初始化 Seagull BarTender 引擎

            engine.Start();

            // 開啟標籤文件
            btFormat = engine.Documents.Open(tempPath, selectedPrinter);
            // 參數說明：標籤路徑，印表機名稱

            // 設定標籤中的欄位值
            btFormat.SubStrings["Importer"].Value = ImportertextBox7.Text;
            btFormat.SubStrings["SerialNumber"].Value = SNtextBox7.Text;
            btFormat.SubStrings["PurchaseOrder"].Value = POtextBox7.Text;
            btFormat.SubStrings["MATERIAL"].Value = MaterialtextBox7.Text;
            btFormat.SubStrings["LOTNumber"].Value = LOTtextBox7.Text;
            btFormat.SubStrings["CNO"].Value = CartontextBox7.Text;
            btFormat.SubStrings["ORIGIN"].Value = OrigintextBox7.Text;
            int padLeft = Convert.ToInt16(MD7_CBX_CNO.Text);
            PrintStart(7, padLeft);

            engine.Stop();

            Properties.Settings.Default.LOTNum7 = LOTtextBox7.Text;
            Properties.Settings.Default.Save();
        }

        private void Resetbutton7_Click(object sender, EventArgs e)
        {
            engine.Stop();
            // 清除所有的 TextBox
            ImportertextBox7.Clear();
            SNtextBox7.Clear();
            POtextBox7.Clear();
            MaterialtextBox7.Clear();
            LOTtextBox7.Clear();
            OrigintextBox7.Clear();
            StarttextBox7.Clear();
            EndtextBox7.Clear();
            PrintQuantitynumericUpDown7.Value = 0;
            AddtextBox7.Text = "1";

            if (string.IsNullOrEmpty(OrigintextBox7.Text))
            {
                // 檢查OrigintextBox是否為空，若是則恢復預設文字和外觀
                OrigintextBox7.Text = "MADE IN";
                OrigintextBox7.ForeColor = SystemColors.GrayText;
                OrigintextBox7.TextAlign = HorizontalAlignment.Left;
            }
        }

        private void SetMultilineMode(System.Windows.Forms.TextBox textBox, bool isMultiline)
        {
            if (isMultiline) // 如果需要多行模式
            {
                textBox.Multiline = true; // 啟用多行
                textBox.Height = (int)(textBox.Font.Height * 3) + 1; // 設置高度以容納三行文本
                textBox.ScrollBars = ScrollBars.Vertical; // 啟用垂直滾動條
                textBox.TextAlign = HorizontalAlignment.Left; // 文本左對齊
            }
            else // 如果不需要多行模式
            {
                textBox.Multiline = false; // 禁用多行
                textBox.Height = 45; // 設置高度為45像素
                textBox.ScrollBars = ScrollBars.None; // 禁用滾動條
                textBox.TextAlign = HorizontalAlignment.Center; // 文本居中對齊
            }
        }

        private void ManufacturertextBox6_Enter(object sender, EventArgs e)
        {
            if (!isUpdating) // 檢查是否正在更新，以避免無窮遞歸
            {
                isUpdating = true; // 設置為正在更新
                SetMultilineMode(ManufacturertextBox6, true); // 啟用多行模式
                isUpdating = false; // 更新完成
            }
        }

        private void ManufacturertextBox6_Leave(object sender, EventArgs e)
        {
            if (!isUpdating) // 檢查是否正在更新
            {
                isUpdating = true; // 設置為正在更新
                SetMultilineMode(ManufacturertextBox6, false); // 關閉多行模式
                OrigintextBox6.Focus(); // 設置焦點到OrigintextBox6
                isUpdating = false; // 更新完成
            }
        }

        private void ImportertextBox7_Enter(object sender, EventArgs e)
        {
            if (!isUpdating) // 檢查是否正在更新
            {
                isUpdating = true; // 設置為正在更新
                SetMultilineMode(ImportertextBox7, true); // 啟用多行模式
                isUpdating = false; // 更新完成
            }
        }

        private void ImportertextBox7_Leave(object sender, EventArgs e)
        {
            if (!isUpdating) // 檢查是否正在更新
            {
                isUpdating = true; // 設置為正在更新
                SetMultilineMode(ImportertextBox7, false); // 關閉多行模式
                SNtextBox7.Focus(); // 設置焦點到SNtextBox7
                isUpdating = false; // 更新完成
            }
        }

        private void richTextBox_Enter(object sender, EventArgs e)
        {
            tabPage8.Focus();
        }

        private void PrintQuantitynumericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            copies = (int)PrintQuantitynumericUpDown1.Value;
        }

        private void PrintQuantitynumericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            copies = (int)PrintQuantitynumericUpDown2.Value;
        }

        private void PrintQuantitynumericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            copies = (int)PrintQuantitynumericUpDown3.Value;
        }

        private void PrintQuantitynumericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            copies = (int)PrintQuantitynumericUpDown4.Value;
        }

        private void PrintQuantitynumericUpDown5_ValueChanged(object sender, EventArgs e)
        {
            copies = (int)PrintQuantitynumericUpDown5.Value;
        }

        private void PrintQuantitynumericUpDown6_ValueChanged(object sender, EventArgs e)
        {
            copies = (int)PrintQuantitynumericUpDown6.Value;
        }

        private void PrintQuantitynumericUpDown7_ValueChanged(object sender, EventArgs e)
        {
            copies = (int)PrintQuantitynumericUpDown7.Value;
        }

        

        private void setTabVisible(ToolStripMenuItem menuItem)
        {
            tabControl.Visible = true;
            
            string currentMenuItemName = menuItem.Text;

            tabControl.TabPages.Clear();

            foreach (TabPage tabPage in _allTabControl)
            {
                if (tabPage.Text == currentMenuItemName)
                {
                    tabControl.TabPages.Add(tabPage);
                }
            }
        }

        private void 模板1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 确保 sender 是一个 ToolStripMenuItem
            if (sender is ToolStripMenuItem menuItem)
            {
                setTabVisible(menuItem);
            }
        }

        private void 模板2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem menuItem)
            {
                setTabVisible(menuItem);
            }
        }

        private void 模板3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem menuItem)
            {
                setTabVisible(menuItem);
            }
        }

        private void 模板4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem menuItem)
            {
                setTabVisible(menuItem);
            }
        }

        private void 模板5ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem menuItem)
            {
                setTabVisible(menuItem);
            }
        }

        private void 模板6ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem menuItem)
            {
                setTabVisible(menuItem);
            }
        }

        private void 模板7ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem menuItem)
            {
                setTabVisible(menuItem);
            }
        }

        private void 列印設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem menuItem)
            {
                setTabVisible(menuItem);
            }
        }
    }
}
