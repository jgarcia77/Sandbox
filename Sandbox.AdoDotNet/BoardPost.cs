using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sandbox.AdoDotNet
{
    public class BoardPost
    {
        public long Id { get; set; }
        public string Message { get; set; }
        public bool IsActive { get; set; }
        public long CreatedMemberId { get; set; }
        public DateTime CreatedDate { get; set; }
        public long UpdatedMemberId { get; set; }
        public DateTime UpdatedDate { get; set; }
        public long BoardTopicId { get; set; }
    }
}
