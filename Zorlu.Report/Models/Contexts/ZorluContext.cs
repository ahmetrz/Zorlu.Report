using Microsoft.EntityFrameworkCore;

namespace Zorlu.Report.Models.Contexts
{
	public class ZorluContext : DbContext
	{

		protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
		{
			if (optionsBuilder.IsConfigured) return;
			//optionsBuilder.UseSqlServer(@"Data Source=(LocalDb)\MSSQLLocalDB;Initial Catalog=Zorlu;Trusted_Connection=True;");
 			optionsBuilder.UseSqlServer(@"Server=tcp:trzahmet.database.windows.net,1433;Initial Catalog=Zorlu;Persist Security Info=False;User ID=zorlu;Password=Ahmet123@;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=100000;");
		}

		public DbSet<Data> Data { get; set; }
		public DbSet<Rapor1> Rapor1 { get; set; }
		public DbSet<IstekIdIstekTipi> IstekIdIstekTipi { get; set; }

		protected override void OnModelCreating(ModelBuilder modelBuilder)
		{
			modelBuilder.Entity<Rapor1>().HasNoKey();
			modelBuilder.Entity<IstekIdIstekTipi>().HasNoKey();
		}


	}
}
