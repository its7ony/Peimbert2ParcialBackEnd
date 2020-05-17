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

        [HttpGet]
        [Route("GetDataPieByDimensionYear/{dim}/{year}/{values}")]
        public HttpResponseMessage GetDataPieByDimension(string dim, string year, string values)
        {
            string MDX_QUERY = string.Empty;
           
                MDX_QUERY = @"
                SELECT 
                   [Dim Tiempo].[Mes].[Mes].AllMembers
                ON COLUMNS,  
                   NONEMPTY (ORDER(
		                 STRTOSET(@Dimension)," +
                                dim + @".CURRENTMEMBER.MEMBER_NAME, ASC
	                ))
   
                ON ROWS  

                FROM [DWH Northwind]

                WHERE " + year + ";";
            
            Debug.Write(MDX_QUERY);

            List<string> dimension = new List<string>();
            List<dynamic> lstVentas = new List<dynamic>();
            List<ObjectGeneric> objectList = new List<ObjectGeneric>();

            dynamic result = new
            {
                tableYear = Regex.Replace(year, "[^0-9]", ""),
                datosDimension = dimension,
                datosVenta = lstVentas
            };

            var valuesArray = values.Split(','); 
            string valoresDimension = string.Empty;
            foreach (var item in valuesArray)
            {
                valoresDimension += "{0}.[" + item + "],";
            }
            valoresDimension = valoresDimension.TrimEnd(',');
            valoresDimension = string.Format(valoresDimension, dim);
            valoresDimension = @"{" + valoresDimension + "}";

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorthwind"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    cmd.Parameters.Add("Dimension", valoresDimension);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        while (dr.Read())
                        {
                            dimension.Add(dr.GetString(0));
                            List<string> auxList = new List<string>();
                            int xValues = dr.FieldCount;
                            for (int i = 1; i < xValues; i++)
                            {
                                string ventaName = string.Empty;
                                try {
                                    auxList.Add(dr.GetString(i));
                                    ventaName = dr.GetString(i);
                                }
                                catch (Exception ex)
                                {
                                    auxList.Add(string.Empty);
                                    ventaName = string.Empty;
                                    Debug.WriteLine(ex.Message);
                                }

                            }


                            dynamic objDinamic = new
                            {
                                descripcion = dr.GetString(0),
                                arrayVenta = auxList
                            };

                            lstVentas.Add(objDinamic);

                    }
                        dr.Close();
                    }
                }
            }

            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }


        [HttpGet]
        [Route("GetDataBarByDimensionMonth/{dim}/{month}/{values}")]
        public HttpResponseMessage GetDataBarByDimensionMonth(string dim, string month, string values)
        {
            string MDX_QUERY = string.Empty;

            MDX_QUERY = @"
                SELECT 
                   [Dim Tiempo].[Año].[Año].AllMembers
                ON COLUMNS,  
                   NONEMPTY (ORDER(
		                 STRTOSET(@Dimension)," +
                            dim + @".CURRENTMEMBER.MEMBER_NAME, ASC
	                ))
   
                ON ROWS  

                FROM [DWH Northwind]

                WHERE " + month + ";";

            Debug.Write(MDX_QUERY);

            List<string> dimension = new List<string>();
            List<dynamic> lstVentas = new List<dynamic>();
            List<ObjectGeneric> objectList = new List<ObjectGeneric>();

            string[] years = { "1996,1997,1998" };

            dynamic result = new
            {
                years = years,
                mes = Regex.Replace(month, "[^0-9]", ""),
                datosDimension = dimension,
                datosVenta = lstVentas
            };

            var valuesArray = values.Split(',');
            string valoresDimension = string.Empty;
            foreach (var item in valuesArray)
            {
                valoresDimension += "{0}.[" + item + "],";
            }
            valoresDimension = valoresDimension.TrimEnd(',');
            valoresDimension = string.Format(valoresDimension, dim);
            valoresDimension = @"{" + valoresDimension + "}";

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorthwind"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    cmd.Parameters.Add("Dimension", valoresDimension);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        while (dr.Read())
                        {
                            dimension.Add(dr.GetString(0));
                            List<string> auxList = new List<string>();
                            int xValues = dr.FieldCount;
                            for (int i = 1; i < xValues; i++)
                            {
                                string ventaName = string.Empty;
                                try
                                {
                                    auxList.Add(dr.GetString(i));
                                    ventaName = dr.GetString(i);
                                }
                                catch (Exception ex)
                                {
                                    auxList.Add(string.Empty);
                                    ventaName = string.Empty;
                                    Debug.WriteLine(ex.Message);
                                }

                            }


                            dynamic objDinamic = new
                            {
                                descripcion = dr.GetString(0),
                                arrayVenta = auxList
                            };

                            lstVentas.Add(objDinamic);

                        }
                        dr.Close();
                    }
                }
            }

            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }

    }

}

