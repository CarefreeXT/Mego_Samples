using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Caredev.Mego;
using Caredev.Mego.DataAnnotations;

namespace ExcelImportExport
{
    partial class Program
    {
        static void Main(string[] args)
        {
            ComplexExcelSample();
            AnonymousExportData();
            AnonymousImportData();
            SimpleImportData();
            SimpleExportData();
        }
    }
}

namespace ExcelImportExport
{
    using Caredev.Mego.Resolve;
    using ExcelImportExport.Simple;
    partial class Program
    {
        static void SimpleImportData()
        {
            using (var context = new ExcelContext("sample.xls"))
            {
                var data = context.Products.ToArray();
            }
        }

        static void SimpleExportData()
        {
            var filename = Guid.NewGuid().ToString() + ".xlsx";
            System.IO.File.WriteAllBytes(filename, Properties.Resources.Empty);
            Random r = new Random();
            var products = Enumerable.Range(0, 1000).Select(i => new Product()
            {
                Id = i,
                Name = "Product " + i.ToString(),
                Category = r.Next(1, 10),
                Code = "P" + i.ToString(),
                IsValid = true
            });

            using (var context = new ExcelContext(filename))
            {
                var operate = context.Database.Manager.CreateTable<Product>();
                context.Executor.Execute(operate);

                context.Products.AddRange(products);
                context.Executor.Execute();
            }
        }

        static void AnonymousImportData()
        {
            using (var context = new ExcelContext("sample.xls"))
            {
                var item = new { Id = 1, Name = "P", IsValid = false };

                var data = context.Set(item, "Products").ToArray();
            }
        }

        static void AnonymousExportData()
        {
            var filename = Guid.NewGuid().ToString() + ".xlsx";
            System.IO.File.WriteAllBytes(filename, Properties.Resources.Empty);
            Random r = new Random();
            var items = Enumerable.Range(0, 1000).Select(i => new
            {
                Id = i,
                Name = "Product " + i.ToString(),
                Category = r.Next(1, 10),
                IsValid = true
            }).ToArray();

            using (var context = new ExcelContext(filename))
            {
                var item = items[0];
                var operate = context.Database.Manager.CreateTable(item.GetType(),
                    DbName.NameOnly("Sheet1$"));
                context.Executor.Execute(operate);

                context.Set(item, "[Sheet1$]").AddRange(items);
                context.Executor.Execute();
            }
        }

    }
}
namespace ExcelImportExport
{
    using Caredev.Mego.Resolve;
    using ExcelImportExport.Complex;
    partial class Program
    {
        static void ComplexExcelSample()
        {
            using (var context = new ComplexContext("sample.xls"))
            {
                var query = from a in context.Orders.Include(a=>a.Details)
                            where a.Id > 4
                            select a;
                var items = query.Take(10).Skip(20).ToArray();
            }
        }
    }
}
namespace ExcelImportExport.Simple
{
    public class AnonymouExcelContext : DbContext
    {
        private const string conStr =
                 @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES'";
        public AnonymouExcelContext(string filename)
            : base(string.Format(conStr, filename), "System.Data.OleDb.Excel")
        {
            this.Configuration.EnableAutoConversionStorageTypes = true;
        }
    }

    public class ExcelContext : DbContext
    {
        private const string conStr =
                   @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES'";
        public ExcelContext(string filename)
            : base(string.Format(conStr, filename), "System.Data.OleDb.Excel")
        {
            this.Configuration.EnableAutoConversionStorageTypes = true;
        }

        public DbSet<Product> Products { get; set; }
    }

    [Table("Products")]
    public class Product
    {
        [Key]
        public int Id { get; set; }

        public string Code { get; set; }

        public string Name { get; set; }

        public int Category { get; set; }

        public bool IsValid { get; set; }
    }
}

namespace ExcelImportExport.Complex
{
    internal class ComplexContext : DbContext
    {
        private const string conStr =
                 @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES'";
        public ComplexContext(string filename)
            : base(string.Format(conStr, filename), "System.Data.OleDb.Excel")
        {
            this.Configuration.EnableAutoConversionStorageTypes = true;
        }
        public DbSet<Order> Orders { get; set; }
        public DbSet<OrderDetail> OrderDetails { get; set; }
        public DbSet<Customer> Customers { get; set; }
        public DbSet<Product> Products { get; set; }
        public DbSet<Warehouse> Warehouses { get; set; }
    }
    [Table("Orders")]
    internal class Order
    {
        [Key]
        public int Id { get; set; }
        public int CustomerId { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime ModifyDate { get; set; }
        public int State { get; set; }
        [ForeignKey("CustomerId", "Id")]
        public virtual Customer Customer { get; set; }
        [InverseProperty("OrderId", "Id")]
        public virtual ICollection<OrderDetail> Details { get; set; }
        [Relationship(typeof(OrderDetail), "OrderId", "Id", "ProductId", "Id")]
        public virtual ICollection<Product> Products { get; set; }
    }

    [Table("OrderDetails")]
    internal class OrderDetail
    {
        [Key]
        public int Id { get; set; }
        public int OrderId { get; set; }
        public int? ProductId { get; set; }
        public string Key { get; set; }
        public decimal Price { get; set; }
        public int Discount { get; set; }
        [ForeignKey("OrderId", "Id")]
        public virtual Order Order { get; set; }
        [ForeignKey("ProductId", "Id")]
        public virtual Product Product { get; set; }
    }
    [Table("Customers")]
    internal class Customer
    {
        [Key]
        public int Id { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public string Zip { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        [InverseProperty("CustomerId", "Id")]
        public virtual ICollection<Order> Orders { get; set; }
    }
    [Table("Products")]
    internal class Product
    {
        [Key]
        public int Id { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public int Category { get; set; }
        public bool IsValid { get; set; }
        public DateTime UpdateDate { get; set; }
        [InverseProperty("OrderId", "Id")]
        public virtual ICollection<Order> Orders { get; set; }
    }
    [Table("Warehouses")]
    internal class Warehouse
    {
        [Key, Column(nameof(Id), Order = 1)]
        public int Id { get; set; }
        [Key, Column(nameof(Number), Order = 2)]
        public int Number { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
    }
}