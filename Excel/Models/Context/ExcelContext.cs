using Excel.Models.Entities;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Excel.Models.Context
{
    public class ExcelContext : DbContext
    {
        public ExcelContext(DbContextOptions options)
            :base(options)
        {

        }

        protected override void OnModelCreating(ModelBuilder builder)
        {
            builder.Entity<Sample_Relation>().HasKey(a => new { a.SampleId, a.RelationId });

            builder.Entity<Sample_Relation>().HasOne(a => a.Sample)
                .WithMany(a => a.Sample_Relation)
                .HasForeignKey(a => a.SampleId);

            builder.Entity<Sample_Relation>().HasOne(a => a.Relation)
                .WithMany(a => a.Sample_Relation)
                .HasForeignKey(a => a.RelationId);
        }

        public DbSet<Relation> Relations { get; set; }
        public DbSet<Sample_Relation> Sample_Relations { get; set; }
        public DbSet<Sample> Samples { get; set; }
    }
}
