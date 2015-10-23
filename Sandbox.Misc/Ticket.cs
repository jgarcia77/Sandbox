using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sandbox.Misc
{
    public class Ticket
    {
        
        public Guid TicketId { get; set; }

        
        public String ExternalId { get; set; }

        
        public String TicketNumber { get; set; }

        
        public Guid? TicketOwnerId { get; set; }

        
        public Guid? CreatedByUserId { get; set; }

        
        public DateTime CreatedDateUtc { get; set; }

        
        public DateTime? OpenedDateUtc { get; set; }

        
        public DateTime? ClosedDateUtc { get; set; }

        
        public String CountryCodeISOAlpha2 { get; set; }

        
        public String ErrorCode { get; set; }

        
        public String ExperienceCode { get; set; }

        
        public String ExperienceCodeDescription { get; set; }

        
        public bool FieldActionMandatory { get; set; }

        
        public String IssueCode { get; set; }

        
        public String IssueCodeDescription { get; set; }

        
        public String LotNumber { get; set; }

        
        public DateTime? PlannedVisitStartDateUtc { get; set; }

        
        public String PurchaseOrderNumber { get; set; }

        
        public bool PotentiallyReportableEvent { get; set; }

        
        public bool ReagentOnly { get; set; }

        
        public object RecordType { get; set; }

        
        public bool RepeatCall { get; set; }

        
        public String ShortDescription { get; set; }

        
        public String Priority { get; set; }

        
        public DateTime DueDateUtc { get; set; }

        
        public String ContractResponseTimeCode { get; set; }


        public object Instrument { get; set; }


        public object PrimaryContact { get; set; }


        public object Customer { get; set; }


        public object Status { get; set; }


        public object DispatchTicketType { get; set; }

        
        public int DurationMinutes { get; set; }

        
        public Guid ProductId { get; set; }

        
        public String ProductName { get; set; }

        
        public String ProductGroupCode { get; set; }

        
        public String ProductListNumber { get; set; }

        
        public bool CanBundleTSB { get; set; }


        public List<object> Skills { get; set; }


        public List<object> BusinessHours { get; set; }
    }
}
