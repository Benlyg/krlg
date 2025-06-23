using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Report.Service
{
    public interface IPrintLabel
    {
        string Print_label(string pathjson, string pathRDLC, string printername);
        string Print_label_PDF(string pathjson, string pathRDLC, string printername);
    }
}
