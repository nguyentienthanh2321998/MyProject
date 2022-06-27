using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Linq.Dynamic.Core;
using System.Threading.Tasks;
using Abp.Application.Services;
using Abp.Authorization;
using Abp.Domain.Repositories;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MyProject.AppServices.Customer;
using MyProject.AppServices.Customer.Dto;
using MyProject.Common;
using MyProject.Common.Files;
using OfficeOpenXml;

namespace MyProject.AppServices;

public class CustomerAppService : ApplicationService, ICustomerAppService
{
    private readonly IRepository<Models.Customer.Customer, long> _customerRepository;

    public CustomerAppService(IRepository<Models.Customer.Customer, long> customerRepository)
    {
        _customerRepository = customerRepository;
    }

    public async Task<CustomerDto> CreateCustomer(CreateCustomerDto input)
    {
        Models.Customer.Customer customer = Models.Customer.Customer.CreateCustomer(input.FullName, input.Birthday, input.Address, input.Phone, input.Email, input.CardId);
        await _customerRepository.InsertAsync(customer);
        await CurrentUnitOfWork.SaveChangesAsync();
        return ObjectMapper.Map<CustomerDto>(customer);
    }

    public async Task<CustomerDto> UpdateCustomer(UpdateCustomerDto input)
    {
        Models.Customer.Customer customer = _customerRepository.GetAll().FirstOrDefault(p => p.Id == input.Id);
        customer.UpdateCustomer(input.FullName, input.Birthday, input.Address, input.Phone, input.Email, input.CardId);
        _customerRepository.Update(customer);
        return ObjectMapper.Map<CustomerDto>(customer);
    }


    public async Task<CustomerDto> DeleteCustomer(long id)
    {
        Models.Customer.Customer customer = _customerRepository.GetAll().FirstOrDefault(p => p.Id == id);
        customer.DeleteCustomer();
        _customerRepository.Update(customer);
        return ObjectMapper.Map<CustomerDto>(customer);
    }



    public FileStreamResult ExportCustomer()
    {
        var customers = _customerRepository.GetAll().ToList();

        var stream = CreatedTemplateExportCustomer(customers);
        return new FileStreamResult(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        {
            FileDownloadName = "Customers"
        };
    }
    private Stream? CreatedTemplateExportCustomer(
        List<Models.Customer.Customer> customers)
    {
        var result = new DataSet();
        var tableUpdatePrice = CreatedTableExportCustomer(customers);
        result.Tables.Add(tableUpdatePrice);
        var stream = CreateWorkbookExportCustomer(result, false);
        ExcelPackage.LicenseContext = LicenseContext.Commercial;

        using var excel = new ExcelPackage(stream);

        return new MemoryStream(excel.GetAsByteArray());
    }
    private static DataTable CreatedTableExportCustomer(
         List<Models.Customer.Customer> customers)
    {
        var table = new DataTable("Customer");
        table.Columns.Add("Id");
        table.Columns.Add("FullName");
        table.Columns.Add("Birthday");
        table.Columns.Add("Address");
        table.Columns.Add("Phone");
        table.Columns.Add("Email");
        table.Columns.Add("CardId");
        table.Columns.Add("CreationTime");
        table.Columns.Add("CreatorUserId");
        table.Columns.Add("LastModificationTime");
        table.Columns.Add("LastModifierUserId");
        table.Columns.Add("IsActive");
        foreach (var row in customers)
        {
            table.Rows.Add(row.Id, row.FullName, row.Birthday, row.Address, row.Phone, row.Email, row.CardId, row.CreationTime, row.CreatorUserId, row.LastModificationTime, row.LastModifierUserId, row.IsActive);
        }


        return table;
    }
    private static Stream CreateWorkbookExportCustomer(
        DataSet dataset,
        bool isCreatedStream = false)
    {
        if (dataset.Tables.Count == 0)
            throw new ArgumentException("DataSet needs to have at least one DataTable", "dataset");

        ExcelPackage.LicenseContext = LicenseContext.Commercial;
        using var excel = new ExcelPackage();

        foreach (DataTable table in dataset.Tables)
        {
            var sheet = excel.Workbook.Worksheets.Add(table.TableName);
            sheet.Protection.IsProtected = true;
            sheet.Cells["A1"].LoadFromDataTable(table, true);
            sheet.Cells.SetStyleDefault(false);

            if (isCreatedStream) continue;
            var columnIndex = 1;
            foreach (DataColumn column in table.Columns)
            {
                var columnCurrent = sheet.Column(columnIndex);
                if (column.DataType == typeof(DateTime))
                    columnCurrent.Style.Numberformat.Format = "MM/dd/yyyy hh:mm:ss AM/PM";
                columnCurrent.Style.Font.Size = 11;
                columnCurrent.AutoFit();

                var header = sheet.Cells[1, columnIndex];
                header.SetStyleDefault();
                header.Style.Font.Size = 11;
                header.Style.WrapText = true;
                switch (header.Value.ToString())
                {
                    case "Id":
                        columnCurrent.Style.Locked = true;
                        //columnCurrent.Hidden = true;
                        break;
                    case "FullName":
                        columnCurrent.Style.Locked = true;
                        //columnCurrent.Hidden = true;
                        break;
                    case "Birthday":
                        columnCurrent.Width = 45;
                        break;
                    case "Address":
                        columnCurrent.Style.Locked = true;
                        columnCurrent.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        header.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        columnCurrent.Width = 51;
                        break;
                    case "Phone":
                        columnCurrent.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        columnCurrent.Style.Locked = true;
                        columnCurrent.Width = 22;
                        break;
                    case "Email":
                        columnCurrent.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        columnCurrent.Style.Locked = true;
                        columnCurrent.Width = 23;
                        break;
                    case "CardId":
                        columnCurrent.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        columnCurrent.Style.Locked = true;
                        columnCurrent.Width = 23;
                        break;
                    case "CreationTime":
                        columnCurrent.Style.Locked = false;
                        columnCurrent.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        header.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        columnCurrent.Width = 18;
                        break;
                    case "CreatorUserId":
                        columnCurrent.Style.Locked = true;
                        header.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        columnCurrent.Width = 29;
                        break;
                    case "LastModificationTime":
                        columnCurrent.Style.Locked = false;
                        columnCurrent.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        header.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        columnCurrent.Width = 13.5;
                        break;
                    case "LastModifierUserId":
                    case "IsActive":
                        columnCurrent.Style.Locked = false;
                        columnCurrent.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        header.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        columnCurrent.Width = 13.5;
                        break;
                };
                columnIndex++;
            }
        }

        var stream = new MemoryStream(excel.GetAsByteArray());

        return stream;
    }

    public async Task<List<CustomerDto>> ImportCustomer(IFormFile input)
    {
        var dataJson = input.OpenReadStream().ExcelToDictionary();
        var tasks = 
            dataJson.FirstOrDefault(item => item.Key.Equals("Customer")).Value
           
            .Select(item =>
            {
                var fullName = item.GetStringValue("FullName");
                var birthday = item.GetStringValue("Birthday");
                var address = item.GetStringValue("Address");
                var phone = item.GetStringValue("Phone");
                var email = item.GetStringValue("Email");
                var cardId = item.GetStringValue("CardId");
                var customer =  Models.Customer.Customer.CreateCustomer(fullName,DateTimeOffset.Parse( birthday), address, phone, email, cardId);
               return   _customerRepository.InsertAsync(customer);

            }).ToList();
       var customers =  await Task.WhenAll(tasks);
            await CurrentUnitOfWork.SaveChangesAsync();
        return ObjectMapper.Map<List<CustomerDto>>(customers);
    } 

    public async Task<List<CustomerDto>> GetCustomers(int? page, int? pageSize)
    {
        var partition = new PartitionHelper(page, pageSize);
       var customers =  _customerRepository.GetAll().Skip(partition.From).Take(partition.PageSize).ToList();
        return ObjectMapper.Map<List<CustomerDto>>(customers);
    }

    public  async Task<CustomerDto> GetCustomer(long id)
    {
        var customer = _customerRepository.FirstOrDefault(p => p.Id == id);
        return ObjectMapper.Map<CustomerDto>(customer);
    }
}