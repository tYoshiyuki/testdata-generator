using ChainingAssertion;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using TestdataGenerator.Library.Extentions;
using Xunit;

namespace TestdataGenerator.Library.Test.Extentions
{
    public class DbContextExtentionTest : IDisposable
    {
        private readonly string _rootPath = Path.Combine(Environment.CurrentDirectory, "Data", "DbContextExtentionTest");
        private readonly TestDbContext _context;

        /// <summary>
        /// Setup
        /// </summary>
        public DbContextExtentionTest()
        {
            _context = new TestDataContextFactory().CreateDbContext(null);
            _context.Database.EnsureCreated(); // DBの初期化
        }

        [Fact]
        public async Task ReadExcelWriteDb_正常系()
        {
            // Arrange
            var path = Path.Combine(_rootPath, "TestData.xlsx");
            _context.Database.ExecuteSqlCommand("truncate table TestObject");

            // Act
            _context.ReadExcelWriteDb(path);
            var tmp1 = (await _context.TestObject.ToListAsync()).Count();
            _context.ReadExcelWriteDb(path);
            var tmp2 = (await _context.TestObject.ToListAsync()).Count();

            // Assert
            (tmp2 - tmp1).Is(3);
            var obj = _context.TestObject.First(_ => _.ColInt == 10);
            obj.ColInt.Is(10);
            obj.ColNullableInt.Is(10);
            obj.ColShort.Is((short)1);
            obj.ColNullableShort.Is((short)1);
            obj.ColLong.Is(10000000000L);
            obj.ColNullableLong.Is(10000000000L);
            obj.ColDecimal.Is((decimal)1.1);
            obj.ColNullableDecimal.Is((decimal)1.1);
            obj.ColDouble.Is(11.1);
            obj.ColNullableDouble.Is(11.1);
            obj.ColDateTime.Is(DateTime.Parse("2000/01/01 00:00:00"));
            obj.ColNullableDateTime.Is(DateTime.Parse("2100/01/01 00:00:00"));
            obj.ColString.Is("AAA");
            obj.ColEnum.Is(TestEnum.Eval1);

            obj = _context.TestObject.First(_ => _.ColInt == 20);
            obj.ColInt.Is(20);
            obj.ColNullableInt.IsNull();
            obj.ColShort.Is((short)2);
            obj.ColNullableShort.IsNull();
            obj.ColLong.Is(20000000000L);
            obj.ColNullableLong.IsNull();
            obj.ColDecimal.Is((decimal)2.2);
            obj.ColNullableDecimal.IsNull();
            obj.ColDouble.Is(22.2);
            obj.ColNullableDouble.IsNull();
            obj.ColDateTime.Is(DateTime.Parse("2000/01/02 00:00:00"));
            obj.ColNullableDateTime.IsNull();
            obj.ColString.Is(string.Empty);
            obj.ColEnum.Is(TestEnum.Eval2);
        }

        [Fact]
        public async Task ReadExcelWriteTable_正常系()
        {
            // Arrange
            var path = Path.Combine(_rootPath, "TestData.xlsx");
            _context.Database.ExecuteSqlCommand("truncate table TestObject");

            // Act
            _context.ReadExcelWriteTable(path, "TestObject");
            var tmp1 = (await _context.TestObject.ToListAsync()).Count();
            _context.ReadExcelWriteTable(path, "TestObject");
            var tmp2 = (await _context.TestObject.ToListAsync()).Count();

            // Assert
            (tmp2 - tmp1).Is(3);
            var obj = _context.TestObject.First(_ => _.ColInt == 10);
            obj.ColInt.Is(10);
            obj.ColNullableInt.Is(10);
            obj.ColShort.Is((short)1);
            obj.ColNullableShort.Is((short)1);
            obj.ColLong.Is(10000000000L);
            obj.ColNullableLong.Is(10000000000L);
            obj.ColDecimal.Is((decimal)1.1);
            obj.ColNullableDecimal.Is((decimal)1.1);
            obj.ColDouble.Is(11.1);
            obj.ColNullableDouble.Is(11.1);
            obj.ColDateTime.Is(DateTime.Parse("2000/01/01 00:00:00"));
            obj.ColNullableDateTime.Is(DateTime.Parse("2100/01/01 00:00:00"));
            obj.ColString.Is("AAA");
            obj.ColEnum.Is(TestEnum.Eval1);

            obj = _context.TestObject.First(_ => _.ColInt == 20);
            obj.ColInt.Is(20);
            obj.ColNullableInt.IsNull();
            obj.ColShort.Is((short)2);
            obj.ColNullableShort.IsNull();
            obj.ColLong.Is(20000000000L);
            obj.ColNullableLong.IsNull();
            obj.ColDecimal.Is((decimal)2.2);
            obj.ColNullableDecimal.IsNull();
            obj.ColDouble.Is(22.2);
            obj.ColNullableDouble.IsNull();
            obj.ColDateTime.Is(DateTime.Parse("2000/01/02 00:00:00"));
            obj.ColNullableDateTime.IsNull();
            obj.ColString.Is(string.Empty);
            obj.ColEnum.Is(TestEnum.Eval2);
        }

        [Fact]
        public async Task ReadExcelReplaceDb_正常系()
        {
            // Arrange
            var path = Path.Combine(_rootPath, "TestData.xlsx");
            _context.Database.ExecuteSqlCommand("truncate table TestObject");

            // Act
            _context.ReadExcelReplaceDb(path);
            var tmp1 = (await _context.TestObject.ToListAsync()).Count();
            _context.ReadExcelReplaceDb(path);
            var tmp2 = (await _context.TestObject.ToListAsync()).Count();

            // Assert
            (tmp2 - tmp1).Is(0);
            var obj = _context.TestObject.First(_ => _.ColInt == 10);
            obj.ColInt.Is(10);
            obj.ColNullableInt.Is(10);
            obj.ColShort.Is((short)1);
            obj.ColNullableShort.Is((short)1);
            obj.ColLong.Is(10000000000L);
            obj.ColNullableLong.Is(10000000000L);
            obj.ColDecimal.Is((decimal)1.1);
            obj.ColNullableDecimal.Is((decimal)1.1);
            obj.ColDouble.Is(11.1);
            obj.ColNullableDouble.Is(11.1);
            obj.ColDateTime.Is(DateTime.Parse("2000/01/01 00:00:00"));
            obj.ColNullableDateTime.Is(DateTime.Parse("2100/01/01 00:00:00"));
            obj.ColString.Is("AAA");
            obj.ColEnum.Is(TestEnum.Eval1);

            obj = _context.TestObject.First(_ => _.ColInt == 20);
            obj.ColInt.Is(20);
            obj.ColNullableInt.IsNull();
            obj.ColShort.Is((short)2);
            obj.ColNullableShort.IsNull();
            obj.ColLong.Is(20000000000L);
            obj.ColNullableLong.IsNull();
            obj.ColDecimal.Is((decimal)2.2);
            obj.ColNullableDecimal.IsNull();
            obj.ColDouble.Is(22.2);
            obj.ColNullableDouble.IsNull();
            obj.ColDateTime.Is(DateTime.Parse("2000/01/02 00:00:00"));
            obj.ColNullableDateTime.IsNull();
            obj.ColString.Is(string.Empty);
            obj.ColEnum.Is(TestEnum.Eval2);
        }

        [Fact]
        public async Task ReadExcelReplaceTable_正常系()
        {
            // Arrange
            var path = Path.Combine(_rootPath, "TestData.xlsx");
            _context.Database.ExecuteSqlCommand("truncate table TestObject");

            // Act
            _context.ReadExcelReplaceTable(path, "TestObject");
            var tmp1 = (await _context.TestObject.ToListAsync()).Count();
            _context.ReadExcelReplaceTable(path, "TestObject");
            var tmp2 = (await _context.TestObject.ToListAsync()).Count();

            // Assert
            (tmp2 - tmp1).Is(0);
            var obj = _context.TestObject.First(_ => _.ColInt == 10);
            obj.ColInt.Is(10);
            obj.ColNullableInt.Is(10);
            obj.ColShort.Is((short)1);
            obj.ColNullableShort.Is((short)1);
            obj.ColLong.Is(10000000000L);
            obj.ColNullableLong.Is(10000000000L);
            obj.ColDecimal.Is((decimal)1.1);
            obj.ColNullableDecimal.Is((decimal)1.1);
            obj.ColDouble.Is(11.1);
            obj.ColNullableDouble.Is(11.1);
            obj.ColDateTime.Is(DateTime.Parse("2000/01/01 00:00:00"));
            obj.ColNullableDateTime.Is(DateTime.Parse("2100/01/01 00:00:00"));
            obj.ColString.Is("AAA");
            obj.ColEnum.Is(TestEnum.Eval1);

            obj = _context.TestObject.First(_ => _.ColInt == 20);
            obj.ColInt.Is(20);
            obj.ColNullableInt.IsNull();
            obj.ColShort.Is((short)2);
            obj.ColNullableShort.IsNull();
            obj.ColLong.Is(20000000000L);
            obj.ColNullableLong.IsNull();
            obj.ColDecimal.Is((decimal)2.2);
            obj.ColNullableDecimal.IsNull();
            obj.ColDouble.Is(22.2);
            obj.ColNullableDouble.IsNull();
            obj.ColDateTime.Is(DateTime.Parse("2000/01/02 00:00:00"));
            obj.ColNullableDateTime.IsNull();
            obj.ColString.Is(string.Empty);
            obj.ColEnum.Is(TestEnum.Eval2);
        }

        /// <summary>
        /// Teardown
        /// </summary>
        public void Dispose()
        {
            _context.Dispose();
        }
    }

    public class TestDbContext : DbContext
    {
        public TestDbContext(DbContextOptions<TestDbContext> options) : base(options) { }

        protected override void OnModelCreating(ModelBuilder modelBuilder) { }

        public DbSet<TestObject> TestObject { get; set; }
    }

    public class TestDataContextFactory : IDesignTimeDbContextFactory<TestDbContext>
    {
        public TestDbContext CreateDbContext(string[] args)
        {
            IConfigurationRoot configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json")
                .Build();

            var builder = new DbContextOptionsBuilder<TestDbContext>();
            var connectionString = configuration.GetConnectionString("DefaultConnection");

            builder.UseSqlServer(connectionString);

            return new TestDbContext(builder.Options);
        }
    }

    public class TestObject
    {
        public int Id { get; set; }
        public int ColInt { get; set; }
        public int? ColNullableInt { get; set; }
        public short ColShort { get; set; }
        public short? ColNullableShort { get; set; }
        public long ColLong { get; set; }
        public long? ColNullableLong { get; set; }
        public decimal ColDecimal { get; set; }
        public decimal? ColNullableDecimal { get; set; }
        public double ColDouble { get; set; }
        public double? ColNullableDouble { get; set; }
        public DateTime ColDateTime { get; set; }
        public DateTime? ColNullableDateTime { get; set; }
        public string ColString { get; set; }
        public TestEnum ColEnum { get; set; }
    }

    public enum TestEnum
    {
        Eval1 = 1, Eval2, Eval3
    }
}
