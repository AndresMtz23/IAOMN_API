using IAOMN_API_REST.Models;
using Microsoft.AnalysisServices.AdomdClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Cors;

namespace IAOMN_API_REST.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*")]
    [RoutePrefix("v1/Proyecto/Northwind")]
    public class NorthwindController : ApiController
    {
        private decimal getDecimalHelper(object value)
        {
            Debug.WriteLine(value);
            try
            {
                return decimal.Parse(value.ToString());
            }
            catch (Exception)
            {
                return 0;
            }
        }

        [HttpGet]
        [Route("Testing")]
        public HttpResponseMessage Testing()
        {
            return Request.CreateResponse(HttpStatusCode.OK, "Funciona!");
        }

        [HttpGet]
        [Route("GetItemsByDimension/{dim}/{order}")]
        public HttpResponseMessage GetItemsByDimension(string dim, string order = "DESC")
        {
           
            string WITH = @"
                WITH 
                SET [OrderDimension] AS 
                NONEMPTY(
                    ORDER(
                        {0}.CHILDREN,
                        {0}.CURRENTMEMBER.MEMBER_NAME, " + order +
                    @")
                )
            ";

            string COLUMNS = @"
                NON EMPTY
                {
                    [Measures].[Hec Ventas Ventas]
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
                    //cmd.Parameters.Add("Dimension", dimension);
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
        [Route("GetDataByDimension/{dim}/{order}")]
        public HttpResponseMessage GetDataByDimension(string dim, string order, string[] values)
        {

            string WITH = @"
            WITH 
                SET [OrderDimension] AS 
                NONEMPTY(
                    ORDER(
			        STRTOSET(@Dimension),
                    [Measures].[Hec Ventas Ventas], DESC
	            )
            )
            ";

            string COLUMNS = @"
                NON EMPTY
                {
                    [Measures].[Hec Ventas Ventas]
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
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + " FROM " + CUBO_NAME;

            Debug.Write(MDX_QUERY);

            List<string> dimension = new List<string>();
            List<decimal> ventas = new List<decimal>();
            List<dynamic> lstTabla = new List<dynamic>();

            dynamic result = new
            {
                datosDimension = dimension,
                datosVenta = ventas,
                datosTabla = lstTabla
            };

            string valoresDimension = string.Empty;
            foreach (var item in values)
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
                            ventas.Add(Math.Round(dr.GetDecimal(1)));

                            dynamic objTabla = new
                            {
                                descripcion = dr.GetString(0),
                                valor = Math.Round(dr.GetDecimal(1))
                            };

                            lstTabla.Add(objTabla);
                        }
                        dr.Close();
                    }
                }
            }

            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }

        [HttpPost]
        [Route("GetDataYearByDimension/{dim}/{year}/{order}")]
        public HttpResponseMessage GetDataYearByDimension(string dim, string year,  string order, Consulta data)
        {
            string WITH = @"
            WITH 
                SET [Items] AS 
                {
                    STRTOSET(@Dimension)
                }
            ";

            string COLUMNS = @"
                NON EMPTY
                {
                    STRTOSET(@Year)
                }
                ON COLUMNS,    
            ";

            string ROWS = @"
                NON EMPTY
                {
                    ORDER(
                        [Items],
                        [Measures].[Hec Ventas Ventas], " + order +
                    @")
                }
                ON ROWS
            ";

            string CUBO_NAME = "[DWH Northwind]";
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + " FROM " + CUBO_NAME;

            Debug.Write(MDX_QUERY);
            List<string> anio = new List<string>();
            List<dynamic> lstTabla = new List<dynamic>();


            string valoresDimension = string.Empty;
            string valoresYear = string.Empty;
            foreach (var item in data.Dimension)
            {
                valoresDimension += "{0}.[" + item + "],";
            }
            valoresDimension = valoresDimension.TrimEnd(',');
            valoresDimension = string.Format(valoresDimension, dim);
            valoresDimension = @"{" + valoresDimension + "}";

            foreach (var item in data.Year)
            {
                valoresYear += "{0}.[" + item + "],";
            }
            valoresYear = valoresYear.TrimEnd(',');
            valoresYear = string.Format(valoresYear, year);
            valoresYear = @"{" + valoresYear + "}";

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorthwind"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    cmd.Parameters.Add("Dimension", valoresDimension);
                    cmd.Parameters.Add("Year", valoresYear);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        while (dr.Read())
                        {
                            switch (dr.FieldCount)
                            {
                                case 2:
                                    dynamic objTabla = new
                                    {
                                        name = dr.GetString(0),
                                        ventas = new decimal[] { getDecimalHelper(dr.GetValue(1))},
                                        total = getDecimalHelper(dr.GetValue(1)),
                                    };
                                    lstTabla.Add(objTabla);
                                    break;
                                case 3:
                                    dynamic objTabla2 = new
                                    {
                                        name = dr.GetString(0),
                                        ventas = new decimal[] { getDecimalHelper(dr.GetValue(1)), getDecimalHelper(dr.GetValue(2))},
                                        total = getDecimalHelper(dr.GetValue(1)) + getDecimalHelper(dr.GetValue(2)),
                                    };
                                    lstTabla.Add(objTabla2);
                                    break;
                                case 4:
                                    dynamic objTabla3 = new
                                    {
                                        name = dr.GetString(0),
                                        ventas = new decimal []{ getDecimalHelper(dr.GetValue(1)), getDecimalHelper(dr.GetValue(2)), getDecimalHelper(dr.GetValue(3)) },
                                        total = getDecimalHelper(dr.GetValue(1)) + getDecimalHelper(dr.GetValue(2)) + getDecimalHelper(dr.GetValue(3))
                                    };
                                    lstTabla.Add(objTabla3);
                                   
                                    break;
                            }
                        }
                        switch (dr.FieldCount)
                        {
                            case 4:
                                anio.Add(dr.GetName(1).Substring(22, 4));
                                anio.Add(dr.GetName(2).Substring(22, 4));
                                anio.Add(dr.GetName(3).Substring(22, 4));
                                break;
                            case 3:
                                anio.Add(dr.GetName(1).Substring(22, 4));
                                anio.Add(dr.GetName(2).Substring(22, 4));
                                break;
                            case 2:
                                anio.Add(dr.GetName(1).Substring(22, 4));
                                break;
                        }
                        dr.Close();
                    }
                }
            }
            dynamic result = new
            {
                anio,
                sales = lstTabla
            };

            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }
        [HttpPost]
        [Route("GetDataYearByDimension/{dim}/{year}/{month}/{order}")]
        public HttpResponseMessage GetDataMonthByDimension(string dim, string year, string month, string order, ConsultaMes data)
        {

            string WITH = @"
            WITH 
                SET [Items] AS 
                {
                    STRTOSET(@Dimension)
                }
            ";

            string COLUMNS = @"
                NON EMPTY
                {
                    STRTOSET(@Year)
                }
                ON COLUMNS,    
            ";

            string ROWS = @"
                NON EMPTY
                {
                    ORDER(
                        [Items],
                        [Measures].[Hec Ventas Ventas], " + order +
                    @")
                }
                *
                {
                    STRTOSET(@Month)
                }
                ON ROWS
            ";

            string CUBO_NAME = "[DWH Northwind]";
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + " FROM " + CUBO_NAME;

            Debug.Write(MDX_QUERY);
            List<string> anio = new List<string>();
            List<dynamic> lstTabla = new List<dynamic>();

            string valoresDimension = string.Empty;
            string valoresYear = string.Empty;
            string valoresMonth = string.Empty;
            foreach (var item in data.Dimension)
            {
                valoresDimension += "{0}.[" + item + "],";
            }
            valoresDimension = valoresDimension.TrimEnd(',');
            valoresDimension = string.Format(valoresDimension, dim);
            valoresDimension = @"{" + valoresDimension + "}";

            foreach (var item in data.Year)
            {
                valoresYear += "{0}.[" + item + "],";
            }
            valoresYear = valoresYear.TrimEnd(',');
            valoresYear = string.Format(valoresYear, year);
            valoresYear = @"{" + valoresYear + "}";

            foreach (var item in data.Month)
            {
                valoresMonth += "{0}.[" + item + "],";
            }
            valoresMonth = valoresMonth.TrimEnd(',');
            valoresMonth = string.Format(valoresMonth, month);
            valoresMonth = @"{" + valoresMonth + "}";

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorthwind"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    cmd.Parameters.Add("Dimension", valoresDimension);
                    cmd.Parameters.Add("Year", valoresYear);
                    cmd.Parameters.Add("Month", valoresMonth);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        while (dr.Read())
                        {
                            switch (dr.FieldCount)
                            {
                                case 2:
                                    dynamic objTabla = new
                                    {
                                        name = dr.GetString(0),
                                        ventas = new decimal[] { getDecimalHelper(dr.GetValue(2)) },
                                        total = getDecimalHelper(dr.GetValue(2)),
                                    };
                                    lstTabla.Add(objTabla);
                                    break;
                                case 3:
                                    dynamic objTabla2 = new
                                    {
                                        name = dr.GetString(0),
                                        ventas = new decimal[] { getDecimalHelper(dr.GetValue(2)), getDecimalHelper(dr.GetValue(3)) },
                                        total = getDecimalHelper(dr.GetValue(2)) + getDecimalHelper(dr.GetValue(3)),
                                    };
                                    lstTabla.Add(objTabla2);
                                    break;
                                case 4:
                                    dynamic objTabla3 = new
                                    {
                                        name = dr.GetString(0),
                                        ventas = new decimal[] { getDecimalHelper(dr.GetValue(2)), getDecimalHelper(dr.GetValue(3)), getDecimalHelper(dr.GetValue(4)) },
                                        total = getDecimalHelper(dr.GetValue(2)) + getDecimalHelper(dr.GetValue(3)) + getDecimalHelper(dr.GetValue(4))
                                    };
                                    lstTabla.Add(objTabla3);

                                    break;
                            }
                        }
                        switch (dr.FieldCount)
                        {
                            case 4:
                                anio.Add(dr.GetName(1).Substring(22, 4));
                                anio.Add(dr.GetName(2).Substring(22, 4));
                                anio.Add(dr.GetName(3).Substring(22, 4));
                                break;
                            case 3:
                                anio.Add(dr.GetName(1).Substring(22, 4));
                                anio.Add(dr.GetName(2).Substring(22, 4));
                                break;
                            case 2:
                                anio.Add(dr.GetName(1).Substring(22, 4));
                                break;
                        }
                        dr.Close();
                    }
                }
            }
            dynamic result = new
            {
                anio,
                sales = lstTabla
            };

            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }
    }


}
