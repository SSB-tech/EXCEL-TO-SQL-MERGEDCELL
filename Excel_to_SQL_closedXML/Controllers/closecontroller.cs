using ClosedXML.Excel;
using Dapper;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel_to_SQL_closedXML.models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System.Data.SqlClient;

namespace Excel_to_SQL_closedXML.Controllers
{
	[Route("api/[controller]")]
	[ApiController]
	public class Closecontroller : ControllerBase
	{
		private readonly IConfiguration _configuration;

		public Closecontroller( IConfiguration configuration)
		{
			_configuration = configuration;
		}

		[HttpPost]
		public ActionResult import(IFormFile file)
		{
			List<close> list = new List<close>();

			MemoryStream stream = new MemoryStream();
			file.CopyTo(stream);

			XLWorkbook workbook = new XLWorkbook(stream); //access/creates a workbook and store the value that you passed through Iformfile file
			IXLWorksheet worksheet = workbook.Worksheet(1); //access/creates a worksheet 
			var row = worksheet.RowsUsed().Count();//counts total number of rows of excel
			var col = worksheet.ColumnsUsed().Count();

			for (int i = 2; i <= row; i++)
			{

				list.Add(new close
				{	//Yo block bhitra worksheet ko value lai Model class ko field ma initialize gareko using loop and creating new object of class in every iteration
					Id = worksheet.Cell(i, 1).GetValue<int>(), //First ma "(int)worksheet.cells[i,1].value;" yo garda null reference error ayo did it this way
					Customercode = worksheet.Cell(i,2).GetValue<int?>(),
					FirstName = worksheet.Cell(i, 3).GetValue<string>(),
					LastName = worksheet.Cell(i, 4).GetValue<string>(),
					gender = worksheet.Cell(i, 5).GetValue<string>(),
					Country = worksheet.Cell(i, 6).GetValue<string?>(),
					Age = worksheet.Cell(i, 7).GetValue<int?>(),
				}
				);
			}

			//to store value of merged cell
			for(int c=0; c < list.Count; c++)
			{
				if (list[c].Customercode == null)
				{
					list[c].Customercode = list[c - 1].Customercode;
				}

				if (list[c].FirstName == "" || list[c].FirstName == null)
				{
					list[c].FirstName = list[c - 1].FirstName;
				}
				
				if (list[c].LastName == "" || list[c].LastName == null)
				{
					list[c].LastName = list[c - 1].LastName;
				}
				
				if (list[c].gender == "" || list[c].gender == null)
				{
					list[c].gender = list[c - 1].gender;
				}
				
				if (list[c].Country == "" || list[c].Country == null)
				{
					list[c].Country = list[c - 1].Country;
				}
				
				if (list[c].Age == null)
				{
					list[c].Age = list[c - 1].Age;
				}

				}


			//Console.WriteLine(list[1].Country); Yo k aucha vanera test garna matra use gareko

			var connection = new SqlConnection(_configuration.GetConnectionString("defaultconnection"));
			connection.ExecuteAsync("insert into closexmltbl (CustomerCode, FirstName, LastName, Gender, Country,Age) values (@CustomerCode, @FirstName, @LastName, @Gender, @Country, @Age)", list);
			return Ok(list);
			
		}
	}
}
