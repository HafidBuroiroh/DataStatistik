using Microsoft.EntityFrameworkCore;
using DataStatistik.Models;


namespace DataStatistik.Data
{
    public class DataContext : DbContext
    {
        public DataContext(DbContextOptions<DataContext> options) : base(options)
        {
        }
        public DbSet<data_statistik> data_statistik { get; set; }
        public DbSet<data_statistik_saved> data_statistik_saved { get; set; }
    }
}
