using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
namespace Decl
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private static List<producer> prod;
        private static List<importer> importLs;
        private static List<supply> supl;
        private static int headerColumnIndex;
        private static bool upLoadToDb = false;
        private static string password_sdf = "7338a7e6-fd3b-49d1-8d90-ddbbc1b39fa1";
        private static SqlCeConnection conn = null;
        private static List<organization> organizations;
        private static string kpp_organization = null;
        private static int priz_period;
        private static int year_otch;


        private void button1_Click(object sender, EventArgs e)
        {
            tabControl1.TabPages.Clear();
            label1.Text = "";
            string path;
            OpenFileDialog Fd = new OpenFileDialog();

            Fd.Title = "Выберете файл с деклорацией";
            Fd.InitialDirectory = @"C:\";
            Fd.Filter = "файлы делораций(xml)|*.xml";
            if (Fd.ShowDialog() == DialogResult.OK)
            {
                prod = null;
                importLs = null;
                supl = null;
                path = Fd.FileName;
                checkBox1.Visible = true;
                checkBox1.Checked = false;
                label2.Visible = true;
                XDocument xdoc = XDocument.Load(path);
                prod = new List<producer>();
                importLs = new List<importer>();
                supl = new List<supply>();
                organizations = new List<organization>();
                List<producer> findProd = new List<producer>();
                string dateDoc = xdoc.Element("Файл").Attribute("ДатаДок").Value;
                string period = "";
                priz_period = Convert.ToInt32(xdoc.Element("Файл").Element("ФормаОтч").Attribute("ПризПериодОтч").Value);
                year_otch = Convert.ToInt32(xdoc.Element("Файл").Element("ФормаОтч").Attribute("ГодПериодОтч").Value);
                switch (xdoc.Element("Файл").Element("ФормаОтч").Attribute("ПризПериодОтч").Value)
                {
                    case "0":
                        period = "4 квартал " + xdoc.Element("Файл").Element("ФормаОтч").Attribute("ГодПериодОтч").Value + " год.";
                        break;
                    case "3":
                        period = "1 квартал " + xdoc.Element("Файл").Element("ФормаОтч").Attribute("ГодПериодОтч").Value + " год.";
                        break;
                    case "6":
                        period = "2 квартал " + xdoc.Element("Файл").Element("ФормаОтч").Attribute("ГодПериодОтч").Value + " год.";
                        break;
                    case "9":
                        period = "3 квартал " + xdoc.Element("Файл").Element("ФормаОтч").Attribute("ГодПериодОтч").Value + " год.";
                        break;
                }
                foreach (XElement prodElement in xdoc.Element("Файл").Element("Справочники").Elements("ПроизводителиИмпортеры"))
                {

                    if (prodElement.Element("ЮЛ") != null)
                    {
                        string producerINN = null;
                        string producerKPP = null;
                        if (prodElement.Element("ЮЛ").Attribute("П000000000005") != null)
                            producerINN = prodElement.Element("ЮЛ").Attribute("П000000000005").Value;
                        else
                            producerINN = "-------";
                        if (prodElement.Element("ЮЛ").Attribute("П000000000006") != null)
                            producerKPP = prodElement.Element("ЮЛ").Attribute("П000000000006").Value;
                        else
                            producerKPP = "---------";
                        prod.Add(new producer()
                        {
                            Id = Convert.ToInt32(prodElement.Attribute("ИДПроизвИмп").Value),
                            Name = prodElement.Attribute("П000000000004").Value.Replace(@"\", ""),
                            INN = producerINN,
                            KPP = producerKPP,
                        });
                    }
                    else
                    {
                        string producerINN = null;
                        string producerKPP = null;
                        if (prodElement.Attribute("П000000000005") != null)
                            producerINN = prodElement.Attribute("П000000000005").Value;
                        else
                            producerINN = "-";
                        if (prodElement.Attribute("П000000000006") != null)
                            producerKPP = prodElement.Attribute("П000000000006").Value;
                        else
                            producerKPP = "-";
                        prod.Add(new producer()
                        {
                            Id = Convert.ToInt32(prodElement.Attribute("ИДПроизвИмп").Value),
                            Name = prodElement.Attribute("П000000000004").Value.Replace(@"\", ""),
                            INN = producerINN,
                            KPP = producerKPP,
                        });
                    }
                }
                foreach (XElement prodElement in xdoc.Element("Файл").Element("Справочники").Elements("Поставщики"))
                {

                    importLs.Add(new importer()
                    {
                        Id = Convert.ToInt32(prodElement.Attribute("ИдПостав").Value),
                        Name = prodElement.Attribute("П000000000007").Value.Replace(@"\", ""),
                        INN = prodElement.Element("ЮЛ").Attribute("П000000000009").Value,
                        KPP = prodElement.Element("ЮЛ").Attribute("П000000000010").Value,
                    });
                }

                int tabPagesCount = 0;
                foreach (XElement moveElement in xdoc.Element("Файл").Element("Документ").Elements("ОбъемОборота"))
                {

                    tabControl1.TabPages.Add("NewTab" + tabPagesCount.ToString());
                    tabControl1.TabPages[tabPagesCount].Name = "NewTab" + tabPagesCount.ToString();
                    Label l1 = new Label();
                    l1.Width = tabPage1.Width - 550;
                    l1.Name = "label";
                    string sobst = moveElement.Attribute("Наим").Value.Replace(@"\", "");
                    if (moveElement.Attribute("КППЮЛ") != null)
                    {
                        sobst += "  /  " + moveElement.Attribute("КППЮЛ").Value;
                    }
                    l1.Text = sobst;
                    l1.Location = new Point(tabControl1.Location.X - 10, tabControl1.Location.Y - 40);
                    tabControl1.TabPages[tabPagesCount].Controls.Add(l1);
                    string check = moveElement.Attribute("НаличиеОборота").Value;
                    if (check.Equals("true"))
                    {
                        string test = moveElement.Element("Оборот").Attribute("П000000000003").Value;
                        tabControl1.TabPages[tabPagesCount].Controls.Add(createTable(Convert.ToInt32(test), moveElement));
                        organizations.Add(new organization()
                        {
                            tabId = tabPagesCount,
                            Name = sobst,
                            availability_of_turnover = true,
                            id_alchol = Convert.ToInt32(test),
                            turnover = moveElement
                        });
                    }


                    else
                    {
                        Label l2 = new Label();
                        l2.Width = tabPage1.Width - 50;
                        l2.Name = "label";
                        l2.Text = "Движения по данному подразделению за " + period + " не было!!";
                        l2.Location = new Point(tabControl1.Location.X - 10, tabControl1.Location.Y - 15);
                        l2.ForeColor = Color.Red;
                        tabControl1.TabPages[tabPagesCount].Controls.Add(l2);
                        organizations.Add(new organization()
                        {
                            tabId = tabPagesCount,
                            Name = sobst,
                            availability_of_turnover = false,
                            id_alchol = 0,
                            turnover = null
                        });
                    }
                    tabPagesCount++;

                }
                tabControl1.SelectedIndex = 0;
                label1.Text = "Дата документа: " + dateDoc + " Отчетность за " + period;
                tabPage1.Parent = null;
            }


        }


        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private DataGridView createTable(int kod, XElement xmlParse)
        {
            DataGridView dgv = new DataGridView();

            if (kod == 500)
            {
                dgv = createPivoTable(xmlParse);
            }
            else
            {
                dgv = createAlcoTable(xmlParse);
            }
            dgv.Location = new Point(tabControl1.Location.X - 10, tabControl1.Location.Y - 15);
            return dgv;
        }
        private DataGridView createPivoTable(XElement xmlParse)
        {
            DataGridView dgv = new DataGridView();
            //dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgv.Columns.Add("col1", "Код вида");
            dgv.Columns.Add("col2", "Производитель");
            dgv.Columns.Add("col3", "Остаток на  начало");
            dgv.Columns.Add("col4", "Закупки у пр-й");
            dgv.Columns.Add("col5", "Закупки - опт");
            dgv.Columns.Add("col6", "закупки -импорт");
            dgv.Columns.Add("col7", "закупки итого");
            dgv.Columns.Add("col8", "Возврат от пок-й");
            dgv.Columns.Add("col9", "Прочее пост-е");
            dgv.Columns.Add("col10", "Поступление всего");
            dgv.Columns.Add("col11", "Розничная продажа");
            dgv.Columns.Add("col12", "Прочий расход");
            dgv.Columns.Add("col13", "возврат поставщику");
            dgv.Columns.Add("col14", "расход всего");
            dgv.Columns.Add("col15", "Остаток на конец периода");
            dgv.Columns.Add("imp", "importer");
            dgv.Columns.Add("idProduce", "Produce");
            dgv.Columns.Add("sob", "sob");
            dgv.Columns["imp"].Visible = false;
            dgv.Columns["idProduce"].Visible = false;
            dgv.Columns["sob"].Visible = false;

            dgv.Name = "dataGridView50";
            dgv.Width = tabControl1.Width - 10;
            dgv.Height = tabControl1.Height - dgv.Rows[0].Height * 5;

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            //dgv.Size = new System.Drawing.Size(300, 300);
            dgv.RowTemplate.ReadOnly = true;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.AllowDrop = false;
            dgv.GridColor = Color.DarkOrange;
            dgv.Anchor = (AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom);
            //dgv.Font = new System.Drawing.Font("Verdana", 9, FontStyle.Bold);
            dgv.ColumnHeadersHeight = 40;
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dgv.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Verdana", 9, FontStyle.Bold);
            dgv.RowHeadersDefaultCellStyle.Font = new System.Drawing.Font("Verdana", 9, FontStyle.Bold);

            ContextMenuStrip m = new ContextMenuStrip();
            ToolStripMenuItem delColumn = new ToolStripMenuItem("Скрыть столбец");
            ToolStripMenuItem resAll = new ToolStripMenuItem("Восстановить все");
            delColumn.Click += delColumn_Click;
            m.Items.AddRange(new[] { delColumn, resAll });
            dgv.ContextMenuStrip = m;

            dgv.ColumnHeaderMouseClick += new DataGridViewCellMouseEventHandler(dgv_ColumnHeaderMouseClick);
            dgv.CellEnter += new DataGridViewCellEventHandler(dgv_CellEnter);

            foreach (XElement moveElement in xmlParse.Elements("Оборот"))
            {
                string kod = moveElement.Attribute("П000000000003").Value;
                foreach (XElement sales in moveElement.Elements("СведПроизвИмпорт"))
                {
                    string imp = "0";
                    string proizv = getPordName(sales.Attribute("ИдПроизвИмп").Value);
                    string sobst = xmlParse.Attribute("Наим").Value;
                    if (xmlParse.Attribute("КППЮЛ") != null)
                    {
                        sobst += "  /  " + xmlParse.Attribute("КППЮЛ").Value;
                    }
                    if (sales.Element("Поставщик") != null)
                    {
                        foreach (XElement supply in sales.Elements("Поставщик"))
                        {
                            imp = supply.Attribute("ИдПоставщика").Value;
                            setSupply(Convert.ToInt32(kod), Convert.ToInt32(sales.Attribute("ИдПроизвИмп").Value), sobst, supply);
                        }
                    }
                    string pn = sales.Element("Движение").Attribute("П100000000006").Value;
                    dgv.Rows.Add(kod, proizv
                      , sales.Element("Движение").Attribute("П100000000006").Value
                      , sales.Element("Движение").Attribute("П100000000007").Value
                      , sales.Element("Движение").Attribute("П100000000008").Value
                      , sales.Element("Движение").Attribute("П100000000009").Value
                      , sales.Element("Движение").Attribute("П100000000010").Value
                      , sales.Element("Движение").Attribute("П100000000011").Value
                      , sales.Element("Движение").Attribute("П100000000012").Value
                      , sales.Element("Движение").Attribute("П100000000013").Value
                      , sales.Element("Движение").Attribute("П100000000014").Value
                      , sales.Element("Движение").Attribute("П100000000015").Value
                      , sales.Element("Движение").Attribute("П100000000016").Value
                      , sales.Element("Движение").Attribute("П100000000017").Value
                      , sales.Element("Движение").Attribute("П100000000018").Value
                      , imp
                      , sales.Attribute("ИдПроизвИмп").Value
                      , sobst);
                }
            }

            return dgv;
        }
        private DataGridView createAlcoTable(XElement xmlParse)
        {
            DataGridView dgv = new DataGridView();
            //dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgv.Columns.Add("col1", "Код вида");
            dgv.Columns.Add("col2", "Производитель");
            dgv.Columns.Add("col3", "Остаток на  начало");
            dgv.Columns.Add("col4", "Поступление от производителей");
            dgv.Columns.Add("col5", "Поступление от  оптовиков");
            dgv.Columns.Add("col6", "Поступление по импорту");
            dgv.Columns.Add("col7", "закупки итого");
            dgv.Columns.Add("col8", "Возврат от покупателейй");
            dgv.Columns.Add("col9", "Прочее поступлениее");
            dgv.Columns.Add("col10", "перемещение");
            dgv.Columns.Add("col11", "Поступление всего");
            dgv.Columns.Add("col12", "расход:продажиа");
            dgv.Columns.Add("col13", "Прочий расход");
            dgv.Columns.Add("col14", "возврат поставщику");
            dgv.Columns.Add("col15", "расход: перемещение");
            dgv.Columns.Add("col16", "расход всего");
            dgv.Columns.Add("col17", "Остаток на конец отч. периода");
            dgv.Columns.Add("col18", "Со старой маркой");
            dgv.Columns.Add("imp", "importer");
            dgv.Columns.Add("idProduce", "Produce");
            dgv.Columns.Add("sob", "sob");
            dgv.Columns["imp"].Visible = false;
            dgv.Columns["idProduce"].Visible = false;
            dgv.Columns["sob"].Visible = false;

            dgv.Name = "dataGridView50";
            dgv.Width = tabControl1.Width - 10;
            dgv.Height = tabControl1.Height - dgv.Rows[0].Height * 5;
            //dgv.Height = tabControl1.Height - 80;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            //dgv.Size = new System.Drawing.Size(300, 300);
            dgv.RowTemplate.ReadOnly = true;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.AllowDrop = false;
            dgv.GridColor = Color.DarkOrange;
            dgv.Anchor = (AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom);
            //dgv.Font = new System.Drawing.Font("Verdana", 9, FontStyle.Bold);
            dgv.ColumnHeadersHeight = 40;
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dgv.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Verdana", 9, FontStyle.Bold);
            dgv.RowHeadersDefaultCellStyle.Font = new System.Drawing.Font("Verdana", 9, FontStyle.Bold);
            ContextMenuStrip m = new ContextMenuStrip();
            ToolStripMenuItem delColumn = new ToolStripMenuItem("Скрыть столбец");
            ToolStripMenuItem resAll = new ToolStripMenuItem("Восстановить все");
            delColumn.Click += delColumn_Click;
            m.Items.AddRange(new[] { delColumn, resAll });
            dgv.ContextMenuStrip = m;

            dgv.ColumnHeaderMouseClick += new DataGridViewCellMouseEventHandler(dgv_ColumnHeaderMouseClick);
            dgv.CellEnter += new DataGridViewCellEventHandler(dgv_CellEnter);

            foreach (XElement moveElement in xmlParse.Elements("Оборот"))
            {
                string kod = moveElement.Attribute("П000000000003").Value;

                foreach (XElement sales in moveElement.Elements("СведПроизвИмпорт"))
                {
                    string imp = "0";
                    string proizv = getPordName(sales.Attribute("ИдПроизвИмп").Value);
                    string sobst = xmlParse.Attribute("Наим").Value;
                    if (xmlParse.Attribute("КППЮЛ") != null)
                    {
                        sobst += "  /  " + xmlParse.Attribute("КППЮЛ").Value;
                    }
                    if (sales.Element("Поставщик") != null)
                    {
                        foreach (XElement supply in sales.Elements("Поставщик"))
                        {
                            imp = supply.Attribute("ИдПоставщика").Value;
                            setSupply(Convert.ToInt32(kod), Convert.ToInt32(sales.Attribute("ИдПроизвИмп").Value), sobst, supply);
                        }
                    }


                    string pn = sales.Element("Движение").Attribute("П100000000006").Value;
                    dgv.Rows.Add(kod, proizv
                      , sales.Element("Движение").Attribute("П100000000006").Value
                      , sales.Element("Движение").Attribute("П100000000007").Value
                      , sales.Element("Движение").Attribute("П100000000008").Value
                      , sales.Element("Движение").Attribute("П100000000009").Value
                      , sales.Element("Движение").Attribute("П100000000010").Value
                      , sales.Element("Движение").Attribute("П100000000011").Value
                      , sales.Element("Движение").Attribute("П100000000012").Value
                      , sales.Element("Движение").Attribute("П100000000013").Value
                      , sales.Element("Движение").Attribute("П100000000014").Value
                      , sales.Element("Движение").Attribute("П100000000015").Value
                      , sales.Element("Движение").Attribute("П100000000016").Value
                      , sales.Element("Движение").Attribute("П100000000017").Value
                      , sales.Element("Движение").Attribute("П100000000018").Value
                      , sales.Element("Движение").Attribute("П100000000019").Value
                      , sales.Element("Движение").Attribute("П100000000020").Value
                      , sales.Element("Движение").Attribute("П100000000021").Value
                      , imp
                      , sales.Attribute("ИдПроизвИмп").Value
                      , sobst
                      );
                }
            }


            return dgv;
        }
        private string getPordName(string idProd)
        {
            string nameProd;
            nameProd = prod.Find(x => x.Id == Convert.ToInt32(idProd)).Name;
            nameProd += ": " + prod.Find(x => x.Id == Convert.ToInt32(idProd)).INN;
            nameProd += "/ " + prod.Find(x => x.Id == Convert.ToInt32(idProd)).KPP;
            return nameProd;
        }
        private string getImpName(string idProd)
        {
            string nameProd;
            nameProd = importLs.Find(x => x.Id == Convert.ToInt32(idProd)).Name;
            nameProd += ": " + importLs.Find(x => x.Id == Convert.ToInt32(idProd)).INN;
            nameProd += "/ " + importLs.Find(x => x.Id == Convert.ToInt32(idProd)).KPP;
            return nameProd;
        }


        private void setSupply(int idGood, int idProd, string sobst, XElement supplyList)
        {
            string idImporter = supplyList.Attribute("ИдПоставщика").Value;
            foreach (XElement doc in supplyList.Elements("Продукция"))
            {

                string date = doc.Attribute("П200000000013").Value;
                string numb = doc.Attribute("П200000000014").Value;
                string quant = doc.Attribute("П200000000016").Value;
                supl.Add(new supply()
                {
                    idProduct = idGood,
                    idProducer = idProd,
                    NameImporter = getImpName(idImporter),
                    sypplyDate = date,
                    numberDocument = numb,
                    quantitProduct = quant,
                    sobst = sobst,
                });

            }
        }

        private void dgv_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            headerColumnIndex = ((DataGridView)sender).Columns[e.ColumnIndex].Index;



            //            ((DataGridView)sender).Columns[e.ColumnIndex].Visible = false;
        }

        private void dgv_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.Rows.Clear();
            if (((DataGridView)sender).Rows[e.RowIndex].Cells["imp"].Value.ToString() != "0")
            {
                string tm = getImpName(((DataGridView)sender).Rows[e.RowIndex].Cells["imp"].Value.ToString());
                int idproduct = Convert.ToInt32(((DataGridView)sender).Rows[e.RowIndex].Cells["col1"].Value);
                int idProducer = Convert.ToInt32(((DataGridView)sender).Rows[e.RowIndex].Cells["idProduce"].Value);
                string sb = Convert.ToString(((DataGridView)sender).Rows[e.RowIndex].Cells["sob"].Value);
                List<supply> findSupply = new List<supply>();
                findSupply = supl.FindAll(item => item.idProduct == idproduct && item.idProducer == idProducer && item.sobst == sb);
                foreach (var ent in findSupply)
                {
                    string[] row = new string[] { ent.NameImporter, ent.numberDocument, ent.sypplyDate, ent.quantitProduct };
                    dataGridView1.Rows.Add(row);
                }

            }

            else
            {
                dataGridView1.Rows[0].Cells[0].Value = ((DataGridView)sender).Rows[e.RowIndex].Cells["col2"].Value.ToString();
                dataGridView1.Rows[0].Cells[0].Style.ForeColor = Color.Red;
            }
        }
        private void delColumn_Click(object sender, EventArgs e)
        {

        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            tabControl1.Height = this.Height - 250;

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                OpenFileDialog db_path = new OpenFileDialog();
                db_path.Title = "Выберете файлбазы данных";
                db_path.InitialDirectory = @"C:\";
                db_path.Filter = "файлы базы данных(sdf)|*.sdf";
                if (db_path.ShowDialog() == DialogResult.OK)
                {

                    try
                    {
                        conn = new SqlCeConnection("Data Source = " + db_path.FileName + ";Password='" + password_sdf + "'");
                        conn.Open();
                        label2.Text = "Подключение к базе данных успешно. Выгрузку можно произвести по каждому подразделению.";
                        label2.ForeColor = Color.Green;
                        upLoadToDb = true;
                        addUploadBtnonTabpage();
                        button2.Visible = true;

                    }
                    catch (Exception ex)
                    {
                        //label2.Text = "База данных выбрана, но к ней не удалось подключится. Выгрузка осуществлена не будет";
                        label2.Text = ex.Message;
                        upLoadToDb = false;
                        checkBox1.Checked = false;
                        deleteUploadBtnFromTabpage();
                        button2.Visible = false;

                    }
                }
            }
            else
            {
                if (conn.State == System.Data.ConnectionState.Open)
                {
                    conn.Close();
                }
                label2.Text = "Выгрузка в базу данных производься не будет ";
                label2.ForeColor = Color.Red;
                upLoadToDb = false;
                deleteUploadBtnFromTabpage();
                button2.Visible = false;
            }
        }



        private void UploadBtn_Click(object sender, EventArgs e)
        {
            Button clickedButton = sender as Button;
            int tabIndex = tabControl1.TabPages.IndexOfKey(clickedButton.Parent.Name);

            int type_id;
            string[] name_and_kpp = organizations.Find(x => x.tabId == tabIndex).Name.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            int id_organization = get_id_organization_from_db(name_and_kpp);
            if (id_organization > 0)
            {
                if (organizations.Find(x => x.tabId == tabIndex).id_alchol == 500) type_id = 12;
                else type_id = 11;
                if (get_dec_header_id(type_id) > 0)
                {
                    insert_decloration_to_db(id_organization, type_id, organizations.Find(x => x.tabId == tabIndex).turnover);
                }
                else
                {
                    DialogResult result = MessageBox.Show($"Для отчетного периода не создана форма {type_id}.Создать её автоматически и продолжить выгрузку?", "Ошибка импорта декларации", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        insert_dec_header(type_id);
                        insert_decloration_to_db(id_organization, type_id, organizations.Find(x => x.tabId == tabIndex).turnover);
                    }
                }
            }
            else
            {
                MessageBox.Show("Данной организации нет в базе данных. Пожалуйста проверьте xml или выгрузите справочники.", "Организация не обнаружена.", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }

        }

        private void insert_dec_header(int type_id)
        {
            string sql;
            SqlCeCommand cmd = new SqlCeCommand();
            cmd.Connection = conn;
            sql = "INSERT INTO DecHeader (type_id,PrizPeriod,PrizFotch,Yearotch,typePK) " +
           $"VALUES ({type_id},{priz_period},1,{year_otch},1)";


            cmd.CommandText = sql;
            int count = cmd.ExecuteNonQuery();
        }

        private void insert_decloration_to_db(int id_organization, int type_id, XElement turnover)
        {
            if (type_id == 11)
            {
                insert_sales_to_dec11(id_organization, turnover);
            }
            else
            {
                insert_sales_to_dec12(id_organization, turnover);
            }
        }

        private void insert_sales_to_dec12(int id_organization, XElement turnover)
        {
            int hid = get_dec_header_id(12);
            foreach (XElement moveElement in turnover.Elements("Оборот"))
            {
                string alcohol_kod = moveElement.Attribute("П000000000003").Value;
                foreach (XElement sales in moveElement.Elements("СведПроизвИмпорт"))
                {
                    string producer_id = get_producer_id_from_db(sales.Attribute("ИдПроизвИмп").Value);
                    string sql;
                    SqlCeCommand cmd = new SqlCeCommand();
                    cmd.Connection = conn;
                    sql = "INSERT INTO Wrk_Contragents (INN,OrgName,OrgType,producer,carrier,RCode,CCode,Area,City,Place,Street,Building,Korp,Flat,Fl_surname,Fl_name,Fl_secname,Fl_address,Foreign_addres,Varnumber)" +
                    $" VALUES ()";


                    cmd.CommandText = sql;
                    int count = cmd.ExecuteNonQuery();
                }
            }
        }

        private void insert_sales_to_dec11(int id_organization, XElement turnover)
        {
            int hid = get_dec_header_id(11);
            foreach (XElement moveElement in turnover.Elements("Оборот"))
            {
                string alcohol_kod = moveElement.Attribute("П000000000003").Value;
                foreach (XElement sales in moveElement.Elements("СведПроизвИмпорт"))
                {
                    int producer_id = Convert.ToInt32(get_producer_id_from_db(sales.Attribute("ИдПроизвИмп").Value));
                    CultureInfo eng = new CultureInfo("en-EN");
                    string sql;
                    SqlCeCommand cmd = new SqlCeCommand();
                    cmd.Connection = conn;
                    sql = "INSERT INTO DecF11 (Hid,vidCode,ProdId,P106,P107,P108,P109,P110,P111,P112,P113,P114,P115" +
                        ",P116,P117,P118,P119,P120,TTYPE,idOrg,P121)" +
                    $" VALUES({hid},'{alcohol_kod}',{producer_id}" +
                    $",{(Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000006").Value, eng).ToString().Replace(",","."))}" +
                    $",{Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000007").Value, eng).ToString().Replace(",",".")}" +
                    $",{Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000008").Value, eng).ToString().Replace(",",".")}" +
                    $",{Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000009").Value, eng).ToString().Replace(",",".")}" +
                    $",{Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000010").Value, eng).ToString().Replace(",",".")}" +
                    $",{Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000011").Value, eng).ToString().Replace(",",".")}" +
                    $",{Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000012").Value, eng).ToString().Replace(",",".")}" +
                    $",{Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000013").Value, eng).ToString().Replace(",",".")}" +
                    $",{Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000014").Value, eng).ToString().Replace(",",".")}" +
                    $",{Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000015").Value, eng).ToString().Replace(",",".")}" +
                    $",{Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000016").Value, eng).ToString().Replace(",",".")}" +
                    $",{Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000017").Value, eng).ToString().Replace(",",".")}" +
                    $",{Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000018").Value, eng).ToString().Replace(",",".")}" +
                    $",{Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000019").Value, eng).ToString().Replace(",",".")}" +
                    $",{Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000020").Value, eng).ToString().Replace(",",".")}" +
                    $",1,'{id_organization.ToString()}',{Convert.ToDecimal(sales.Element("Движение").Attribute("П100000000021").Value, eng).ToString().Replace(",",".")})";
                    cmd.CommandText = sql;
                    int count = cmd.ExecuteNonQuery();
                    if (sales.Elements("Поставщик") != null)
                        foreach (XElement arrival in sales.Elements("Поставщик"))
                        {
                            insert_arrival_to_dec11(hid, alcohol_kod, producer_id, id_organization, arrival);
                        }
                }
            }
        }

        private void insert_arrival_to_dec11(int hid, string alcohol_kod, int producer_id, int id_organization, XElement arrival)
        {
            CultureInfo eng = new CultureInfo("en-US");
            int importer_id = Convert.ToInt32(arrival.Attribute("ИдПоставщика").Value);
            int license_importer_id = Convert.ToInt32(arrival.Attribute("ИдЛицензии").Value);
            SqlCeCommand cmd = new SqlCeCommand();
            cmd.Connection = conn;
// insert foreach 
foreach
            string sql = "INSERT INTO DecF11 (Hid,vidCode,ProdId,idPost,idLic,P213,P214,P215,P216,TTYPE,idOrg)" +
                $" VALUES({hid},'{alcohol_kod}',{producer_id},{importer_id},{license_importer_id}" +
                $",'{arrival.Element("Продукция").Attribute("П200000000013").Value}'" +
                $",'{arrival.Element("Продукция").Attribute("П200000000014").Value}'" +
                $",'{arrival.Element("Продукция").Attribute("П200000000015").Value}'" +
                $",{Convert.ToDecimal(arrival.Element("Продукция").Attribute("П200000000016").Value, eng).ToString().Replace(",",".")}" +
                $",2,{id_organization})";
            cmd.CommandText = sql;
            int count = cmd.ExecuteNonQuery();
        }

        private string get_producer_id_from_db(string idProd)
        {
            string producer_id = null;
            producer prd = prod.Find(x => x.Id == Convert.ToInt32(idProd));
            producer_id = check_producer_in_db(prd).ToString();
            return producer_id;
        }

        private int get_dec_header_id(int type_id)
        {
            string sql = null;
            int id = 0;
            sql = $"SELECT id FROM DecHeader WHERE type_id ={type_id} and PrizPeriod={priz_period} and Yearotch={year_otch}";


            SqlCeCommand cmd = new SqlCeCommand();
            cmd.Connection = conn;
            cmd.CommandText = sql;
            using (SqlCeDataReader reader = cmd.ExecuteResultSet(ResultSetOptions.Scrollable))
            {
                while (reader.Read())
                {
                    id = reader.GetInt32(0);
                }
                return id;
            }
        }
        private int get_id_organization_from_db(string[] name_and_kpp)
        {
            int id_organization = 0;
            string sql = null; SqlCeCommand cmd = new SqlCeCommand();
            cmd.Connection = conn;
            if (name_and_kpp.Length > 1)
            {
                sql = "SELECT id FROM Wrk_org Where OrgName='" + name_and_kpp[0].Trim() + "' and INN='" + name_and_kpp[1].Trim() + "'";
                cmd.CommandText = sql;

            }
            else
            {
                sql = "SELECT id FROM Wrk_org Where OrgName=@org_name";
                cmd.CommandText = sql;
                cmd.Parameters.Add(new SqlCeParameter("org_name", name_and_kpp[0]));
            }
            using (SqlCeDataReader reader = cmd.ExecuteResultSet(ResultSetOptions.Scrollable))
            {
                while (reader.Read())
                {
                    id_organization = reader.GetInt32(0);
                }
            }

            return id_organization;
        }

        private void get_list_of_period_from_db()
        {
            string sql = "Select * From DecHeader";
            SqlCeCommand cmd = new SqlCeCommand();
            cmd.Connection = conn;
            cmd.CommandText = sql;
            using (SqlCeDataReader reader = cmd.ExecuteResultSet(ResultSetOptions.Scrollable))
            {
                if (reader.HasRows) sql = "sdsdsd";
            }
        }
        private void addUploadBtnonTabpage()
        {
            foreach (TabPage tpb in tabControl1.TabPages)
            {
                int tab_index = tabControl1.TabPages.IndexOfKey(tpb.Name);
                if (check_availability_turnover(tab_index))
                {
                    Button uploadBtn = new Button();
                    uploadBtn.Name = "uploadBtn" + tabControl1.SelectedIndex.ToString();
                    uploadBtn.Text = "Загрузить";
                    uploadBtn.Location = new Point(tabControl1.Location.X + 530, tabControl1.Location.Y - 45);
                    tpb.Controls.Add(uploadBtn);
                    uploadBtn.Click += UploadBtn_Click;
                }
            }
        }

        private bool check_availability_turnover(int selectedIndex)
        {
            return organizations.Find(x => x.tabId == selectedIndex).availability_of_turnover;
        }
        private void deleteUploadBtnFromTabpage()
        {
            foreach (TabPage tpb in tabControl1.TabPages)
            {
                tpb.Controls.OfType<Button>().ToList().ForEach(btn => btn.Dispose());
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            import_producer_to_db();
            import_importer_to_db();
            import_organization_to_db();
            MessageBox.Show("Cправочники успешно выгружены.", "Выгрузка завершена", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

        }

        private string get_kpp_organization()
        {
            string sql = "Select kpp From wrk_org where OrgType=1";
            string kpp = "";
            SqlCeCommand cmd = new SqlCeCommand();
            cmd.Connection = conn;
            cmd.CommandText = sql;
            using (SqlCeDataReader reader = cmd.ExecuteResultSet(ResultSetOptions.Scrollable))
            {
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        kpp = reader.GetString(0);
                    }


                }
                return kpp;
            }
        }

        private void import_organization_to_db()
        {
            foreach (organization imported_organization in organizations)
            {
                if (check_organization_in_db(imported_organization) == false)
                {
                    insert_organization_to_db(imported_organization);
                }
            }
        }

        private void insert_organization_to_db(organization imported_organization)
        {
            string[] name_and_kpp = imported_organization.Name.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string sql;
            SqlCeCommand cmd = new SqlCeCommand();
            cmd.Connection = conn;
            if (name_and_kpp.Length < 2)
            {
                sql = "INSERT INTO wrk_org (KPP,OrgName,Head_id) VALUES " +
                    "('" + get_kpp_organization() + "','"
                    + name_and_kpp[0].Trim() + "',1)";
            }
            else
            {
                sql = "INSERT INTO wrk_org (INN,KPP,OrgName,Head_id) VALUES " +
                                   "('" + name_and_kpp[1].Trim() + "','"
                                   + get_kpp_organization() + "','"
                                   + name_and_kpp[0].Trim() + "',1)";
            }

            cmd.CommandText = sql;
            int count = cmd.ExecuteNonQuery();
        }
        private bool check_organization_in_db(organization imported_organization)
        {

            string[] name_and_kpp = imported_organization.Name.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string sql = null;
            if (name_and_kpp.Length < 2)
            {
                sql = "SELECT * FROM wrk_org WHERE OrgName='" + name_and_kpp[0].Trim() + "'";
            }
            else
            {
                sql = "SELECT * FROM wrk_org WHERE OrgName='" + name_and_kpp[0].Trim() + "' and INN='" + name_and_kpp[1].Trim() + "'";
            }
            SqlCeCommand cmd = new SqlCeCommand();
            cmd.Connection = conn;
            cmd.CommandText = sql;
            using (SqlCeDataReader reader = cmd.ExecuteResultSet(ResultSetOptions.Scrollable))
            {
                if (reader.HasRows) return true;
                else return false;
            }
        }

        private void import_importer_to_db()
        {
            foreach (importer imported_importer in importLs)
            {
                if (check_importer_in_db(imported_importer) == false)
                {
                    insert_importer_to_db(imported_importer);
                }
            }
        }
        private void insert_importer_to_db(importer imported_importer)
        {
            string sql;
            SqlCeCommand cmd = new SqlCeCommand();
            cmd.Connection = conn;
            if (imported_importer.KPP.Equals("-"))
            {
                sql = "INSERT INTO Wrk_Contragents (INN,OrgName,OrgType,producer,carrier,RCode,CCode,Area,City,Place,Street,Building,Korp,Flat,Fl_surname,Fl_name,Fl_secname,Fl_address,Foreign_addres,Varnumber) VALUES " +
                    "('" + imported_importer.INN + "','"
                    + imported_importer.Name + "',"
                    + 1 + ","
                    + "'false',"
                    + "'true','01','','','','','','','','','','','','','','')";
            }
            else
            {
                sql = "INSERT INTO Wrk_Contragents (INN,KPP,OrgName,OrgType,producer,carrier,RCode,CCode,Area,City,Place,Street,Building,Korp,Flat,Fl_surname,Fl_name,Fl_secname,Fl_address,Foreign_addres,Varnumber) VALUES " +
      "('" + imported_importer.INN + "','"
      + imported_importer.KPP + "','"
      + imported_importer.Name + "',"
      + 1 + ","
      + "'false',"
      + "'true','01','','','','','','','','','','','','','','')";
            }

            cmd.CommandText = sql;
            int count = cmd.ExecuteNonQuery();
        }
        private bool check_importer_in_db(importer imported_importer)
        {
            string sql = null;
            if (imported_importer.KPP.Equals("-"))
            {
                sql = "SELECT * FROM Wrk_Contragents WHERE INN='" + imported_importer.INN + "'";
            }
            else
            {
                sql = "SELECT * FROM Wrk_Contragents WHERE INN='" + imported_importer.INN + "' and KPP='" + imported_importer.KPP + "'";
            }
            SqlCeCommand cmd = new SqlCeCommand();
            cmd.Connection = conn;
            cmd.CommandText = sql;
            using (SqlCeDataReader reader = cmd.ExecuteResultSet(ResultSetOptions.Scrollable))
            {
                if (reader.HasRows) return true;
                else return false;
            }
        }

        private void import_producer_to_db()
        {
            foreach (producer imported_producer in prod)
            {
                if (check_producer_in_db(imported_producer) == 0)
                {
                    insert_producer_to_db(imported_producer);
                }
            }
        }
        private void insert_producer_to_db(producer imported_producer)
        {
            string sql;
            SqlCeCommand cmd = new SqlCeCommand();
            cmd.Connection = conn;
            if (imported_producer.KPP.Equals("-"))
            {
                sql = "INSERT INTO Wrk_Contragents (INN,OrgName,OrgType,producer,carrier,RCode,CCode,Area,City,Place,Street,Building,Korp,Flat,Fl_surname,Fl_name,Fl_secname,Fl_address,Foreign_addres,Varnumber) VALUES " +
                    "('" + imported_producer.INN + "','"
                    + imported_producer.Name + "',"
                    + 1 + ","
                    + "'true',"
                    + "'false','01','','','','','','','','','','','','','','')";
            }
            else
            {
                sql = "INSERT INTO Wrk_Contragents (INN,KPP,OrgName,OrgType,producer,carrier,RCode,CCode,Area,City,Place,Street,Building,Korp,Flat,Fl_surname,Fl_name,Fl_secname,Fl_address,Foreign_addres,Varnumber) VALUES " +
      "('" + imported_producer.INN + "','"
      + imported_producer.KPP + "','"
      + imported_producer.Name + "',"
      + 1 + ","
      + "'true',"
      + "'false','01','','','','','','','','','','','','','','')";
            }

            cmd.CommandText = sql;
            int count = cmd.ExecuteNonQuery();
        }
        private int check_producer_in_db(producer imported_producer)
        {
            int id_producer = 0;
            string sql = null;
            if (imported_producer.KPP.Equals("-"))
            {
                sql = "SELECT Id FROM Wrk_Contragents WHERE INN='" + imported_producer.INN + "'";
            }
            else
            {
                sql = "SELECT Id FROM Wrk_Contragents WHERE INN='" + imported_producer.INN + "' and KPP='" + imported_producer.KPP + "'";
            }
            SqlCeCommand cmd = new SqlCeCommand();
            cmd.Connection = conn;
            cmd.CommandText = sql;
            using (SqlCeDataReader reader = cmd.ExecuteResultSet(ResultSetOptions.Scrollable))
            {
                if (reader.HasRows)
                    while (reader.Read())
                    {
                        id_producer = reader.GetInt32(0);
                    }
            }
            return id_producer;
        }
    }

    public class producer
    {
        public string Name { get; set; }
        public int Id { get; set; }
        public string INN { get; set; }
        public string KPP { get; set; }
    }
    public class importer
    {
        public string Name { get; set; }
        public int Id { get; set; }
        public string INN { get; set; }
        public string KPP { get; set; }
    }
    public class supply
    {
        public int idProduct { get; set; }
        public int idProducer { get; set; }
        public string NameImporter { get; set; }
        public string sypplyDate { get; set; }
        public string numberDocument { get; set; }
        public string quantitProduct { get; set; }
        public string sobst { get; set; }
    }
    public class movement
    {
        public string kodPr { get; set; }
        public string bnBalance { get; set; }
        public string prod { get; set; }
        public string buys { get; set; }
        public string buyOpt { get; set; }
        public string buyImp { get; set; }
        public string byuAll { get; set; }
        public string retIn { get; set; }
        public string inPr { get; set; }
        public string allPr { get; set; }
        public string sale { get; set; }
        public string saleIn { get; set; }
        public string retOut { get; set; }
        public string saleAll { get; set; }
        public string endBalance { get; set; }
    }
    public class organization
    {
        public int tabId { get; set; }
        public string Name { get; set; }
        public int id_alchol { get; set; }
        public bool availability_of_turnover { get; set; }
        public XElement turnover { get; set; }

    }


}
