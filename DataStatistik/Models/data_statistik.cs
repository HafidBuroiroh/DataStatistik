using Microsoft.EntityFrameworkCore.Metadata.Internal;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DataStatistik.Models
{
    public class data_statistik
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public DateTime Period { get; set; }
        public int MemberCode { get; set; }
        public string? MemberName { get; set; }
        public string? Province { get; set; }
        public string? City { get; set; }
        public int Frequency { get; set; }
        public int Volume { get; set; }
        public int Value { get; set; }
    }
}
