using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sandbox.Misc
{
    public class Instrument
    {
        
        public Guid InstrumentId { get; set; }

        
        public String SerialNumber { get; set; }

        
        public String InstrumentState { get; set; }

        
        public Guid? LinkInstrumentId { get; set; }

        
        public Guid ProductId { get; set; }

        
        public String ProductName { get; set; }

        
        public String ProductGroupCode { get; set; }

        
        public String ProductListNumber { get; set; }

        
        public Guid? UserId1 { get; set; }

        
        public String UserName1 { get; set; }

        
        public Guid? UserId2 { get; set; }

        
        public String UserName2 { get; set; }

        
        public Guid? UserId3 { get; set; }

        
        public String UserName3 { get; set; }

        
        public DateTime? WarrantyStartDateUtc { get; set; }

        
        public DateTime? WarrantyEndDateUtc { get; set; }

        
        public DateTime? ServicingGroupSite { get; set; }

        
        public object Customer { get; set; }


        public object PrimaryContact { get; set; }
    }
}
