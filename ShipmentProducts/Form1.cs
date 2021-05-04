using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShipmentProducts
{
    public partial class Form1 : Form
    {
        List<BaseClass> ListBase;
        BaseClass Object;  

        public Form1()
        {
            InitializeComponent();
            foreach (var item in new List<string>() {"GS Приемники" })            
                ListType.Items.Add(item);
            GetStatus();
        }

        void GetStatus()
        {
            using (var fas = new FASEntities())
            {
               var list = fas.FAS_Shipped_Status_TB.Select(c => c.StatusName).ToList();
                list.Add("");
                CBStatus.DataSource = list;
                CBStatus.Text = "";
            }
        }

        private void BT_OK_Click(object sender, EventArgs e)
        {
            ClickType();
        }

        private void OkBT_Click(object sender, EventArgs e)
        {
            //if (GetStID() == 3)
            //    return;
            Object.OpenShippedSetting(int.Parse(GridReport[4, GridReport.CurrentRow.Index].Value.ToString()));
            GetGridReport();
        }
                
        private void GridReport_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (GetStID() == 3)
            //    return;

            Object.OpenShippedSetting(int.Parse(GridReport[4, GridReport.CurrentRow.Index].Value.ToString()));
            GetGridReport();
        }

        private void ListType_DoubleClick(object sender, EventArgs e)
        {
            ClickType();
        }

        void ClickType()
        {
            if (ListType.SelectedIndex == -1)
                return;

            ListBase = new List<BaseClass>() { new GS(), new Contract() };
            Object = ListBase[ListType.SelectedIndex];
            Object.control = this;

            Object.OpenType(GB);            
            CBClient.DataSource = Object.GetListClients(); //Получаем список Заказчиков 
            CBClient.Text = "";
            CBModel.DataSource = null;
            NUM.Value = 0;
        }

        private void CBClient_SelectionChangeCommitted(object sender, EventArgs e) //Заказчик
        {
            //if (string.IsNullOrEmpty(CBClient.Text))
            //    return;
            Object.GetCustomerID(CBClient.Text);
            CBModel.DataSource = Object.GetModelList();
            CBModel.Text = "";
        }

        private void CBModel_SelectionChangeCommitted(object sender, EventArgs e) //Модель
        {
            Object.GetModelID(CBModel.Text);
            CBLOT.DataSource = Object.GetLots();
            CBLOT.Text = "";
        }

        private void CBLOT_SelectionChangeCommitted(object sender, EventArgs e) //Заказ
        {
            Object.GetLotID(CBLOT.Text);
        }

        int GetStID()
        {
            using (var fas = new FASEntities())
            {
                return fas.FAS_Shipped_Status_TB.Where(c => c.StatusName == CBStatus.Text).Select(c => c.ID).FirstOrDefault();
            }
        }

        private void CBStatus_SelectionChangeCommitted(object sender, EventArgs e) //Тип отгрузки
        {
            GetGridReport();
        }

        void GetGridReport()
        {
            GridReport.Visible = true;
            LoadGrid.Loadgrid(GridReport, $@"  use fas SELECT count Кол_во,date Дата,concat(LOTCode, ' | ', FULL_LOT_Code) Лот, st.StatusName Статус, s.ID
              FROM [FAS].[dbo].[FAS_Shipped_Table] s
              left join FAS_GS_LOTs g on s.LOTID = g.LOTID
              left join FAS_Shipped_Status_TB st on s.Status = st.ID
              where StatusName is not null and s.Status = {GetStID()}   order by date desc ");
            GridReport.Columns[4].Visible = false;
        }

        private void BTOk_Click(object sender, EventArgs e)
        {
            if (CBClient.Text == "") { 
                MessageBox.Show("Не заполнены поля"); return;
            }
            if (CBModel.Text == "") {
                MessageBox.Show("Не заполнены поля"); return;
            }
            if (CBLOT.Text == ""){
                MessageBox.Show("Не заполнены поля"); return;
            }
            if (NUM.Value == 0){
                MessageBox.Show("Не заполнены поля"); return;
            }

            Object.GetTable((int)NUM.Value);
            GC.Collect();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            GC.Collect();            
        }

       
    }
}
