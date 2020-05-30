using Microsoft.AnalysisServices.AdomdClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Cors;
using WebAPI2P.Models;

namespace WebAPI2P.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*")]
    [RoutePrefix("v1/Analysis/Northwind")]
    public class NorthwindController : ApiController
    {

        [HttpGet]
        [Route("GetItemsByDimension/{dim}")]
        public HttpResponseMessage GetItemsByDimension(string dim)
        {
            string WITH = @"
                WITH 
                SET [OrderDimension] AS 
                NONEMPTY(
                    ORDER(
                        {0}.CHILDREN,
                        {0}.CURRENTMEMBER.MEMBER_NAME, ASC)
                )";

            string COLUMNS = @"
                NON EMPTY
                {
                   [Measures].[Ventas]
                }
                ON COLUMNS,    
            ";

            string ROWS = @"
                NON EMPTY
                {
                    [OrderDimension]
                }
                ON ROWS
            ";

            string CUBO_NAME = "[DWH Northwind]";
            WITH = string.Format(WITH, dim);
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + " FROM " + CUBO_NAME;

            Debug.Write(MDX_QUERY);

            List<string> dimension = new List<string>();

            dynamic result = new
            {
                datosDimension = dimension
            };

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorthwind"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        while (dr.Read())
                        {
                            dimension.Add(dr.GetString(0));
                        }
                        dr.Close();
                    }
                }
            }

            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }

        [HttpPost]
        [Route("GetDataByDimension/{dim}")]
        public HttpResponseMessage GetDataByDimension(string dim, [FromBody] dynamic values)
        {
            string WITH = @"
            WITH 
                SET [OrderDimension] AS 
                NONEMPTY(
                    ORDER(
			        STRTOSET(@Dimension),
                    [Measures].[Ventas], DESC
	            )
            )
            ";

            string COLUMNS = @"
                NON EMPTY
                {
                    [Measures].[Ventas]
                }
                ON COLUMNS,    
            ";

            string ROWS = @"
                NON EMPTY
                {
                    ([OrderDimension], STRTOSET(@Years), STRTOSET(@Months))
                }
                ON ROWS
            ";

            string CUBO_NAME = "[DWH Northwind]";
  
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + " FROM " + CUBO_NAME;

            Debug.Write(MDX_QUERY);

            List<string> dimension = new List<string>();
            List<string> years = new List<string>();
            List<string> months = new List<string>();
            List<decimal> sales = new List<decimal>();
            List<dynamic> tableList = new List<dynamic>();

            dynamic result = new
            {
                dimensionData = dimension,
                yearsData = years,
                monthsData = months,
                salesData = sales,
                tableData = tableList
            };

            string dimensionValues = string.Empty;
            Console.WriteLine(values);
            foreach (var item in values.clients)
            {
                dimensionValues += "{0}.[" + item + "],";
            }
            dimensionValues = dimensionValues.TrimEnd(',');
            dimensionValues = string.Format(dimensionValues, dim);
            dimensionValues = @"{" + dimensionValues + "}";

            string yearsValues = string.Empty;
            foreach (var item in values.years)
            {
                yearsValues += "[Dim Tiempo].[Año].[" + item + "],";
            }
            yearsValues = yearsValues.TrimEnd(',');
            yearsValues = @"{" + yearsValues + "}";

            string monthsValues = string.Empty;
            foreach (var item in values.months)
            {
                monthsValues += "[Dim Tiempo].[Mes].[" + item + "],";
            }
            monthsValues = monthsValues.TrimEnd(',');
            monthsValues = @"{" + monthsValues + "}";

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorthwind"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    cmd.Parameters.Add("Dimension", dimensionValues);
                    cmd.Parameters.Add("Years", yearsValues);
                    cmd.Parameters.Add("Months", monthsValues);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        while (dr.Read())
                        {
                            dimension.Add(dr.GetString(0));
                            years.Add(dr.GetString(1));
                            months.Add(dr.GetString(2));
                            sales.Add(Math.Round(dr.GetDecimal(3)));

                            dynamic objTable = new
                            {
                                description = dr.GetString(0),
                                years = dr.GetString(1),
                                months = dr.GetString(2),
                                valor = Math.Round(dr.GetDecimal(3))
                            };

                            tableList.Add(objTable);
                        }
                        dr.Close();
                    }
                }
            }

            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }

        public string formatMonth(int value)
        {
            string month = string.Empty;
            switch (value)
            {
                case 1:
                    month = "Enero";
                    break;

                case 2:
                    month = "Febrero";
                    break;


                case 3:
                    month = "Marzo";
                    break;


                case 4:
                    month = "Abril";
                    break;


                case 5:
                    month = "Mayo";
                    break;


                case 6:
                    month = "Junio";
                    break;

                case 7:
                    month = "Julio";
                    break;

                case 8:
                    month = "Agosto";
                    break;

                case 9:
                    month = "Septiembre";
                    break;

                case 10:
                    month = "Octubre";
                    break;

                case 11:
                    month = "Novimiembre";
                    break;

                case 12:
                    month = "Diciembre";
                    break;
            }
            return month;
        }
    }

}

