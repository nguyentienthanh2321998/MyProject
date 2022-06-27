using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Abp.Application.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MyProject.AppServices.Customer.Dto;

namespace MyProject.AppServices.Customer;

public interface ICustomerAppService : IApplicationService
{
    Task<CustomerDto> CreateCustomer(CreateCustomerDto input);
    Task<List<CustomerDto>> GetCustomers(int? page , int? pageSize);
    Task<CustomerDto> GetCustomer(long id );
    Task<CustomerDto> UpdateCustomer(UpdateCustomerDto input);
    Task<CustomerDto> DeleteCustomer(long id);
    Task<List<CustomerDto>> ImportCustomer(IFormFile input);
    FileStreamResult ExportCustomer();


}