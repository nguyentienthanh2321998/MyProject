using System;

namespace MyProject.AppServices.Customer.Dto
{
    public class CreateCustomerDto
    {
        public string FullName { get; set; }
        public DateTimeOffset Birthday { get; set; }
        public string Address { get; set; }
        public string Phone { set; get; }
        public string Email { set; get; }
        public string CardId { set; get; }
    }
}