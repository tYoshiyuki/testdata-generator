using Microsoft.EntityFrameworkCore;
using TestdataGenerator.Models;

namespace TestdataGenerator.Data
{
    public class TestDataContext : DbContext
    {
        public TestDataContext(DbContextOptions<TestDataContext> options) : base(options) { }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Character>(entity =>
            {
                entity.Property(e => e.Level).IsRequired();
                entity.Property(e => e.Name).IsRequired();
                entity.Property(e => e.Hp).IsRequired();
                entity.Property(e => e.Mp).IsRequired();
            });

            modelBuilder.Entity<Item>(entity =>
            {
                entity.Property(e => e.Name).IsRequired();
                entity.Property(e => e.CreateDay).IsRequired();
            });

        }
        public DbSet<Character> Character { get; set; }

        public DbSet<Item> Item { get; set; }
    }
}
