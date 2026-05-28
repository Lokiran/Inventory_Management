using Microsoft.EntityFrameworkCore;
using EmployeeAssetManagement.Api.Models;

namespace EmployeeAssetManagement.Api.Data
{
    public class InventoryContext : DbContext
    {
        public InventoryContext(DbContextOptions<InventoryContext> options) : base(options)
        {
        }

        public DbSet<Employee> Employees { get; set; }
        public DbSet<Asset> Assets { get; set; }
        public DbSet<AssetRequest> AssetRequests { get; set; }
        public DbSet<Incident> Incidents { get; set; }
        public DbSet<AssetAssignment> AssetAssignments { get; set; }
        public DbSet<AuditLog> AuditLogs { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            // Employee Configuration
            modelBuilder.Entity<Employee>()
                .HasKey(e => e.EmployeeId);
            modelBuilder.Entity<Employee>()
                .HasIndex(e => e.Email)
                .IsUnique();

            // Asset Configuration
            modelBuilder.Entity<Asset>()
                .HasKey(a => a.AssetId);
            modelBuilder.Entity<Asset>()
                .HasIndex(a => a.SerialNumber)
                .IsUnique();

            // AssetRequest Configuration
            modelBuilder.Entity<AssetRequest>()
                .HasKey(ar => ar.RequestId);
            modelBuilder.Entity<AssetRequest>()
                .HasOne(ar => ar.Employee)
                .WithMany(e => e.AssetRequests)
                .HasForeignKey(ar => ar.EmployeeId)
                .OnDelete(DeleteBehavior.NoAction);
            modelBuilder.Entity<AssetRequest>()
                .HasOne(ar => ar.Asset)
                .WithMany(a => a.AssetRequests)
                .HasForeignKey(ar => ar.AssetId)
                .OnDelete(DeleteBehavior.SetNull);

            // Incident Configuration
            modelBuilder.Entity<Incident>()
                .HasKey(i => i.IncidentId);
            modelBuilder.Entity<Incident>()
                .HasOne(i => i.Employee)
                .WithMany(e => e.Incidents)
                .HasForeignKey(i => i.EmployeeId)
                .OnDelete(DeleteBehavior.NoAction);
            modelBuilder.Entity<Incident>()
                .HasOne(i => i.Asset)
                .WithMany(a => a.Incidents)
                .HasForeignKey(i => i.AssetId)
                .OnDelete(DeleteBehavior.NoAction);

            // AssetAssignment Configuration
            modelBuilder.Entity<AssetAssignment>()
                .HasKey(aa => aa.AssignmentId);
            modelBuilder.Entity<AssetAssignment>()
                .HasOne(aa => aa.Employee)
                .WithMany(e => e.AssetAssignments)
                .HasForeignKey(aa => aa.EmployeeId)
                .OnDelete(DeleteBehavior.NoAction);
            modelBuilder.Entity<AssetAssignment>()
                .HasOne(aa => aa.Asset)
                .WithMany(a => a.AssetAssignments)
                .HasForeignKey(aa => aa.AssetId)
                .OnDelete(DeleteBehavior.NoAction);

            // AuditLog Configuration
            modelBuilder.Entity<AuditLog>()
                .HasKey(al => al.LogId);
            modelBuilder.Entity<AuditLog>()
                .HasOne(al => al.Employee)
                .WithMany()
                .HasForeignKey(al => al.EmployeeId)
                .OnDelete(DeleteBehavior.NoAction);
        }
    }
}
