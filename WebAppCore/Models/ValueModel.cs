namespace WebAppCore.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    
    
    public class ValueModel
    {
        public int Id { get; set; }

        [Required(ErrorMessage = "Name is required")]
        public string Name { get; set; }

        public DateTime CreatedOn { get; set; }
    }
}
