using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShipmentProducts
{
    class Contract : BaseClass
    {
        int CustomerID { get; set; }

        public override List<string> GetListClients() 
          {
              using (var fas = new FASEntities())
              {                
                  var list = fas.Contract_LOT.Where(c => c.IsActive == true & c.Ct_OperLog != null).Select(c => fas.CT_Сustomers.Where(v => v.ID == c.СustomersID).Select(b => b.СustomerName).FirstOrDefault()).Distinct().ToList();
                  list.Add("");
                  return list;
              }
          }

        public override void RemoveShipped(int id)
        {

        }

        public override List<string> GetModelList()
          {
                using (var fas = new FASEntities())
                {
                    var list = fas.FAS_Models.Where(c => fas.Contract_LOT.Where(v => v.IsActive == true && v.СustomersID == CustomerID).Select(b => b.ModelID).ToList().Contains(c.ModelID)).Select(c => c.ModelName).ToList();
                    list.Add("");
                    return list;
            }
          }
          public override void GetCustomerID(string name)
          {
              using (var fas = new FASEntities())
              {
                  CustomerID =  fas.CT_Сustomers.Where(c => c.СustomerName == name).Select(c => c.ID).FirstOrDefault();
              }
          }
        public override void GetShippedTable(DataGridView Grid)
        {
           
        }

        public override void GetTable(int count)
        {
            
        }

        public override void GetLotID(string name)
        {
           
        }

        public override List<string> GetLots()
        {
            return new List<string>();
        }



        //public override int GetLotID(string Name)
        //{
        //  using (var fas = new FASEntities())
        //  {
        //      return fas.Contract_LOT.Where(c => c.LOTCode + " | " + c.FullLOTCode == Name).Select(c => c.ID).FirstOrDefault();
        //  }

        //}
    }
}
