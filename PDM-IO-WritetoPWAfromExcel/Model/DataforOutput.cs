using System;
using System.Collections.Generic;
using System.Text;

namespace PDM_IO_WriteToPWA_ReadXLSX.Model
{
    public class DataforOutput
    {
        public List<dataExportItem> data { get; set; }
    }

    public class dataExportItem
    {
        public string ProjectID { get; set; }
        public string TaskID { get; set; }
        public float TaskCountEstimatedCost { get; set; }
    }

}
