//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан по шаблону.
//
//     Изменения, вносимые в этот файл вручную, могут привести к непредвиденной работе приложения.
//     Изменения, вносимые в этот файл вручную, будут перезаписаны при повторном создании кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace ShipmentProducts
{
    using System;
    using System.Collections.Generic;
    
    public partial class FAS_PackingGS
    {
        public int SerialNumber { get; set; }
        public byte LiterID { get; set; }
        public short LiterIndex { get; set; }
        public short PalletNum { get; set; }
        public short BoxNum { get; set; }
        public short UnitNum { get; set; }
        public System.DateTime PackingDate { get; set; }
        public short PackingByID { get; set; }
        public Nullable<short> LOTID { get; set; }
    
        public virtual FAS_SerialNumbers FAS_SerialNumbers { get; set; }
    }
}
