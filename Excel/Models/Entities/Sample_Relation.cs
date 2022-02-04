using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Excel.Models.Entities
{
    public class Sample_Relation
    {
        public int SampleId { get; set; }
        public int RelationId { get; set; }

        public virtual Sample Sample { get; set; }
        public virtual Relation Relation { get; set; }
    }
}
