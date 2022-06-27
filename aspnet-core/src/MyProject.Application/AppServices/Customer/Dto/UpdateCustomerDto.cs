using System;
using Abp.AutoMapper;

namespace MyProject.AppServices.Customer.Dto
{
    public class UpdateCustomerDto
    {
        public long Id { get; set; }
        public string FullName { get; set; }
        public DateTimeOffset Birthday { get; set; }
        public string Address { get; set; }
        public  string Phone { set; get; }
        public  string Email { set; get; }
        public  string CardId { set; get; }

    }
}

