using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UserMaintenance.Entities
{
    public class User
    {
        public Guid ID { get; } = Guid.NewGuid();

        public string FullName
        {
            get; set;
        }
    }
}

