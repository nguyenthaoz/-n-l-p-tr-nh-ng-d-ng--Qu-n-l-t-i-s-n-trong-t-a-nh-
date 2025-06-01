using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhanMemQuanLyTaiSan.model
{
    public class ThanhPhan
    {
        public string Name { get; set; }
        public string CreatedBy { get; set; }
        public string CreatedOn { get; set; }
        public string TypeName { get; set; }
        public string Space { get; set; }
        public string Description { get; set; }
        public string ExtSystem { get; set; }
        public string ExtObject { get; set; }
        public string ExtIdentifier { get; set; }
        public string SerialNumber { get; set; }
        public string InstallationDate { get; set; }
        public string WarrantyStartDate { get; set; }
        public string TagNumber { get; set; }
        public string AssetIdentifier { get; set; }
        public string Area { get; set; }
        public string Length { get; set; }
    }
}
