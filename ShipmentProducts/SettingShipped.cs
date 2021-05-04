using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShipmentProducts
{
    public partial class SettingShipped : Form
    {

        int id;
        BaseClass BaseClass;
        public SettingShipped(int id, BaseClass baseClass)
        {
            InitializeComponent();
            this.id = id;
            this.BaseClass = baseClass;
        }

        string GetLotName()
        {
            using (var fas = new FASEntities())
            {
                return fas.FAS_Shipped_Table.Where(c => c.ID == id).Select(c => fas.FAS_GS_LOTs.Where(v => v.LOTID == c.LOTID).Select(b => b.FULL_LOT_Code).FirstOrDefault()).FirstOrDefault();
            }
        }

        private void button1_Click(object sender, EventArgs e) //Удалить отгрузку
        {
            var Result = MessageBox.Show($"Вы точно уверены, что хотите удалить отгрузку по лоту - \n {GetLotName()}  ??", "Предупреждение!",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
            if (Result == DialogResult.Yes) { 
                BaseClass.RemoveShipped(id);
                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e) //Отгрузка
        {
            var Result = MessageBox.Show($"Вы подтвержаете, что выбранные данные по лоту - \n {GetLotName()} отгружены??", "Предупреждение!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (Result == DialogResult.Yes) { 
                BaseClass.SetShipped(id);
                this.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
