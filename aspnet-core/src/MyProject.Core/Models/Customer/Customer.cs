using System;
using System.ComponentModel.DataAnnotations;

namespace MyProject.Models.Customer
{
    public class Customer : AuditedSoftDeletableEntity<long> 
    {
        private Customer()
        {

        }
        [Required] public string FullName { get; private set; }
        [Required] public DateTimeOffset Birthday { get; private set; }
        public string Address { get; set; }
        [Required] public string Phone { private set; get; }
        [Required] public string Email { private set; get; }
        [Required] public string CardId { private  set; get; }

        public static Customer CreateCustomer(string fullName , DateTimeOffset birthday , string address , string phone , string email , string cardId)
        {
            return new Customer() {FullName =fullName, Birthday = birthday, IsActive = true,Address  = address, Phone = phone, Email = email,CardId = cardId};
            
        }

        public void UpdateCustomer(string fullName, DateTimeOffset birthday, string address, string phone, string email, string cardId)
        {
            FullName = fullName;
            Birthday = birthday;
            Address = address;
            Phone = phone;
            Email = email;
            CardId = cardId;

        }

        public void DeleteCustomer()
        {
            IsActive = false;
        }
    }
}