using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace appsvc_fnc_dev_bulkuserimport
{
    public class BulkInfo
    {
        public string listID { get; set; }
        public string siteID { get; set; }
    }

    public class UsersList
    {
        public string Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string DepartmentEmail { get; set; }
        public string WorkEmail { get; set; }
    }

}
