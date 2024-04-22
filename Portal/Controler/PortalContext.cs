using Model;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Controler
{
    public class PortalContext:DbContext
    {
        public PortalContext() : base("Data Source=192.168.3.180;Initial Catalog=Portal;Persist Security Info=True;User ID=PCServicesAdministrator;Password=PCS2020")
        {

        }

        public DbSet<User> Users { get; set; }
        public DbSet<History> Histories { get; set; }
        public DbSet<TransAction> TransActions { get; set; }
        public DbSet<Memory> Memories { get; set; }
        public DbSet<Customer> Customers { get; set; }
    }
}
