using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace Excel.Models.ViewModels
{
    public class GeneralViewModel
    {
        public int RelationId { get; set; }

        [Required]
        public string CarName { get; set; }

        [Required]
        public string Color { get; set; }
        public string Price { get; set; }

        [Required]
        public DateTime MadeOn { get; set; }

        public int SampleId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }

        [Required(ErrorMessage = "this is required !")]
        public string PhoneNumber { get; set; }

        [Required]
        [DataType(DataType.EmailAddress)]
        public string Email { get; set; }
        public int Count { get; set; }

        [MinLength(10), MaxLength(100)]
        public string Address { get; set; }
        public bool IsIS { get; set; }
    }
}
