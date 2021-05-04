using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShipmentProducts
{
    public abstract class BaseClass
    {
        public BaseClass()
        {
           
        }

        public Control control { get; set; }

        public static int LotID { get; set; }

        public  int ModelID { get; set; }

        public int StatusShippID { get; set; }
        public DataGridView TempTable { get; set; }        
        public abstract List<string> GetListClients();
        public abstract List<string> GetModelList();
        public abstract void GetCustomerID(string name);

        public abstract void GetLotID(string name);

        public abstract List<string> GetLots();

        public abstract void GetTable(int count);

        public abstract void GetShippedTable(DataGridView Grid);

        public abstract void RemoveShipped(int id);

        public void SetShipped(int id)
        {
            using (var fas = new FASEntities())
            {
                var sh = fas.FAS_Shipped_Table.Where(c => c.ID == id).FirstOrDefault();
                sh.Status = 3;
                fas.SaveChanges();
            }
        }


        public void OpenShippedSetting(int id)
        {
            SettingShipped st = new SettingShipped(id,this);
            st.ShowDialog();
        }
        

        public void GetModelID(string name)
        {
            using (var fas = new FASEntities())
            {
                ModelID = fas.FAS_Models.Where(c => c.ModelName == name).Select(c => c.ModelID).FirstOrDefault();
            }
        }

        public void OpenType(Control control)
        {
            control.Location = new Point(354, 38);
            control.Size = new Size(1012, 380);
            control.Visible = true;
        }

    }
}
