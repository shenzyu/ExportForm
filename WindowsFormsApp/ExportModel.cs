using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp
{
    public class ExportModel
    {
        public virtual string Id { get; set; }
        public virtual string Name { get; set; }
        public virtual string Spec { get; set; }
        public virtual string Unit { get; set; }
        public virtual string Num { get; set; }
        public virtual string Memo { get; set; }
    }
}
