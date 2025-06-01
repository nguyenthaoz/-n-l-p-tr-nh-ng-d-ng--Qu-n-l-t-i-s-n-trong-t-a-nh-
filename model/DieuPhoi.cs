using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhanMemQuanLyTaiSan.model
{
    public class DieuPhoi
    {
        public string Name { get; set; }
        public string CreatedBy { get; set; }
        public string CreatedOn { get; set; }
        public string Category { get; set; }
        public string SheetName { get; set; }
        public string RowName { get; set; }
        public string CoordinateXAxis { get; set; }
        public string CoordinateYAxis { get; set; }
        public string CoordinateZAxis { get; set; }
        public string ExtSystem { get; set; }
        public string ExtObject { get; set; }
        public string ExtIdentifier { get; set; }
        public string ClockwiseRotation { get; set; }
        public string ElevationalRotation { get; set; }
        public string YawRotation { get; set; }
    }
}
