using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using homeBudget.Models;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;

namespace homeBudget.Data
{
    public class TransactionContext :DbContext
    {
        public TransactionContext(DbContextOptions<TransactionContext> options) : base(options)
        {
            
        }
        
        public DbSet<SubCategory> SubCategories { get; set; }
        public DbSet<BankAccount> BankAccounts { get; set; }
        public DbSet<Transaction> Transactions { get; set; }
        public DbSet<KeeWord> KeeWords { get; set; }
        public DbSet<TransactionsKeewords> TransactionsKeewordses { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            //modelBuilder.Entity<SubCategory>().ToTable("SubCategory");
            //modelBuilder.Entity<BankAccount>().ToTable("BankAccount");
            //modelBuilder.Entity<Transaction>().ToTable("Transaction");
            //modelBuilder.Entity<KeeWord>().ToTable("KeeWord");
            //modelBuilder.Entity<TransactionsKeewords>().ToTable("TransactionsKeewords");
        }

    }
}
