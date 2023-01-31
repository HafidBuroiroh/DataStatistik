using DataStatistik.Models;
using System;
using System.Linq;

namespace DataStatistik.Data
{
    public class DataDummy
    {
        public static void Initialize(DataContext context)
        {
            context.Database.EnsureCreated();
            if (context.data_statistik.Any())
            {
                return;   // DB has been seeded
            }
            var dummystatistik = new data_statistik[]
            {
                new data_statistik{Date = DateTime.Parse("2023-01-01"), Period = DateTime.Parse("2023-01-01"), MemberCode = 0001, MemberName = "Zein Mukaffi", Province = "DKI Jakarta", City = "Jakarta Selatan", Frequency = 190290, Volume = 109201, Value = 1},
                new data_statistik{Date = DateTime.Parse("2023-01-02"), Period = DateTime.Parse("2023-01-02"), MemberCode = 0002, MemberName = "Reza Ardhana", Province = "Jawa Barat", City = "Depok", Frequency = 190120, Volume = 107201, Value = 2},
                new data_statistik{Date = DateTime.Parse("2023-01-03"), Period = DateTime.Parse("2023-01-03"), MemberCode = 0003, MemberName = "Rendy Afriatama", Province = "Sumatera Selatan", City = "Lampung", Frequency = 20290, Volume = 32201, Value = 3},
                new data_statistik{Date = DateTime.Parse("2023-01-04"), Period = DateTime.Parse("2023-01-04"), MemberCode = 0004, MemberName = "Muhammad Alvin", Province = "Jawa Tengah", City = "Semarang", Frequency = 28260, Volume = 028201, Value = 4},
                new data_statistik{Date = DateTime.Parse("2023-01-05"), Period = DateTime.Parse("2023-01-05"), MemberCode = 0005, MemberName = "Muhammad Airil", Province = "Jawa Timur", City = "Malang", Frequency = 17291, Volume = 542691, Value = 5},
                new data_statistik{Date = DateTime.Parse("2023-01-06"), Period = DateTime.Parse("2023-01-06"), MemberCode = 0006, MemberName = "Azka Nathan", Province = "DI Yogyakarta", City = "Yogyakarta", Frequency = 586467, Volume = 323109, Value = 6},
            };
            foreach (data_statistik s in dummystatistik)
            {
                context.data_statistik.Add(s);
            }
            context.SaveChanges();
        }
    }
}
