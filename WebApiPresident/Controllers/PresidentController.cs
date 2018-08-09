using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web.Http;
using WebApiPresident.Models;

namespace WebApiPresident.Controllers
{
    public class PresidentController : ApiController
    {
        // GET: api/President
        public IEnumerable<President> Get()
        {
            return ReadExcel();
        }

        [HttpGet]
        public IEnumerable<President> GetPresidentByName(string name)
        {
            return ReadExcel().Where(p => p.Name.ToLowerInvariant().Contains(name)).ToList();
        }

        [HttpGet]
        [Route("api/President/GetPresidentSortByBirthday/{sortBy}")]
        public IEnumerable<President> GetPresidentSortByBirthday(string sortBy)
        {
            var lstPresident = ReadExcel();

            if (sortBy.Equals("desc"))
            {
                return lstPresident.OrderByDescending(p => p.Birthday).ToList();
            }

            return lstPresident.OrderBy(p => p.Birthday).ToList();
        }

        [HttpGet]
        [Route("api/President/GetPresidentSortByDeathday/{sortBy}")]
        public IEnumerable<President> GetPresidentSortByDeathday(string sortBy)
        {
            var lstPresident = ReadExcel();

            if (sortBy.Equals("desc"))
            {
                return lstPresident.OrderByDescending(p => p.Deathday).ToList();
            }

            return lstPresident.OrderBy(p => p.Deathday == DateTime.MinValue? DateTime.MaxValue: p.Deathday).ThenBy(p => p.Deathday).ToList();

        }

        // GET: api/President/5
        public string Get(int id)
        {
            return "value";
        }

        // POST: api/President
        public void Post([FromBody]string value)
        {
        }

        // PUT: api/President/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE: api/President/5
        public void Delete(int id)
        {
        }

        private OleDbConnection OpenConnection(string path)
        {
            OleDbConnection oledbConn = null;
            try
            {
                if (Path.GetExtension(path) == ".xls")
                {
                    oledbConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + path +
                    "; Extended Properties= \"Excel 8.0;HDR=Yes;IMEX=2\"");
                }
                else if (Path.GetExtension(path) == ".xlsx")
                {
                    oledbConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" +
                    path + "; Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';");
                }
                oledbConn.Open();
            }
            catch (Exception ex)
            {
                //Error
            }
            return oledbConn;
        }

        private IList<President> ReadExcel()
        {
            string path = System.Web.HttpContext.Current.Server.MapPath("~/Document/Presidents.xlsx");

            IList <President> objPresidentInfo = new List<President>();
            try
            {
                OleDbConnection oledbConn = OpenConnection(path);
                if (oledbConn.State == ConnectionState.Open)
                {
                    objPresidentInfo = ExtractPresidentExcel(oledbConn);
                    oledbConn.Close();
                }
            }
            catch (Exception ex)
            {
                // Error
            }
            return objPresidentInfo;
        }

        private IList<President> ExtractPresidentExcel(OleDbConnection oledbConn)
        {
            OleDbCommand cmd = new OleDbCommand();
            OleDbDataAdapter oleda = new OleDbDataAdapter();
            DataSet dsPresidentInfo = new DataSet();

            cmd.Connection = oledbConn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM [Sheet1$]"; //Excel Sheet Name
            oleda = new OleDbDataAdapter(cmd);
            oleda.Fill(dsPresidentInfo, "Sheet1");

            var dsPresidentInfoList = dsPresidentInfo.Tables[0].AsEnumerable().Select(s => new President
            {
                Name = Convert.ToString(s["President"] != DBNull.Value ? s["President"] : ""),
                Birthday = Convert.ToDateTime(s["Birthday"] != DBNull.Value ? s["Birthday"] : ""),
                Birthplace = Convert.ToString(s["Birthplace"] != DBNull.Value ? s["Birthplace"] : ""),
                Deathday = Convert.ToDateTime(s["Death day"] != DBNull.Value ? s["Death day"] : null),
                Deathplace = Convert.ToString(s["Death place"] != DBNull.Value ? s["Death place"] : ""),
            }).ToList();

            return dsPresidentInfoList;
        }

        private string GetValue(SpreadsheetDocument doc, Cell cell)
        {
            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
            }
            return value;
        }
    }
}
