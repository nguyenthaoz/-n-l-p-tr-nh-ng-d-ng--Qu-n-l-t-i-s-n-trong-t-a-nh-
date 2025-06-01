using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhanMemQuanLyTaiSan.model
{
    public class ThuocTinh
    {
        public string Name { get; set; }
        public string CreatedBy { get; set; }
        public string CreatedOn { get; set; }
        public string Category { get; set; }
        public string SheetName { get; set; }
        public string RowName { get; set; }
        public string Value { get; set; }
        public string Unit { get; set; }
        public string ExtSystem { get; set; }
        public string ExtObject { get; set; }
        public string ExtIdentifier { get; set; }
        public string Description { get; set; }
        public string AllowedValues { get; set; }
    }
}
