using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace Excel.Models.Entities
{
    public class Relation
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public string CarName { get; set; }

        [Required]
        public string Color { get; set; }
        public string Price { get; set; }

        [Required]
        public DateTime MadeOn { get; set; }

        public virtual List<Sample_Relation> Sample_Relation { get; set; }
    }
}
