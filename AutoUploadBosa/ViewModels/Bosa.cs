using System;
using System.Collections.Generic;
using System.IO;
using System.Data.OracleClient;
using System.Configuration;
using System.Net;
using System.Threading.Tasks;
using Syncfusion.XlsIO;
using System.Data.SqlClient;

namespace AutoUploadBosa.ViewModels
{
    public class Bosa
    {

#pragma warning disable
        private OracleConnection OpenDatabase()
        {
            var gConnectionScc = new OracleConnection(ConfigurationManager.ConnectionStrings["Gemtek_SCC"].ConnectionString);
            gConnectionScc.Open();
            return gConnectionScc;
        }
        public async Task WriteDataBosaScc(string filePath)
        {
            string sheetName = "";
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                var gConnectionScc = OpenDatabase();
                try
                {
                    IApplication app = excelEngine.Excel;
                    app.DefaultVersion = ExcelVersion.Excel2016;
                    IWorkbook workbook = app.Workbooks.Open(filePath, ExcelOpenType.Automatic);
                    IWorksheet worksheet = workbook.Worksheets[0];
                    int lastRow = worksheet.UsedRange.LastRow;
                    List<string> sheetNames = await GetSheetName(filePath);
                    foreach (string item in sheetNames)
                    {
                        sheetName = item;
                    }
                    string gPN = sheetName;
                    int totalColumns = await GetTotalColumns(filePath);
                    ProcessDataScc(totalColumns, worksheet, lastRow, gPN, gConnectionScc);
                }
                finally
                {
                    gConnectionScc.Close();
                }
            }
        }
        private void ProcessDataScc(int totalColumns, IWorksheet worksheet, int lastRow, string gPN, OracleConnection gConnectionScc)
        {
            switch (totalColumns)
            {
                case 21:
                    ProcessTwentyColumns(worksheet, lastRow, gPN, gConnectionScc);
                    break;
                case 25:
                    ProcessTwentyFourColumns(worksheet, lastRow, gPN, gConnectionScc);
                    break;
                case 26:
                    ProcessTwentyFiveColumns(worksheet, lastRow, gPN, gConnectionScc);
                    break;
                case 42:
                    ProcessFortyOneColumns(worksheet, lastRow, gPN, gConnectionScc);
                    break;
                default:
                    break;
            }
        }



        /*The writing excel file has 20 column*/
        private void ProcessTwentyColumns(IWorksheet worksheet, int lastRow, string gPN, OracleConnection gConnectionScc)
        {
            for (int row = 6; row <= lastRow; row++)
            {
                if (CheckCellsNotEmpty(worksheet, row, new int[] { 1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 12, 17, 18, 19, 13, 14, 15 }))
                {
                    string gCustData = CreateDataTwentyColumns(worksheet, row);
                    ExecuteDatabase(worksheet, row, gPN, gCustData, gConnectionScc);
                }
                //iCount += 1;
            }
        }
        private string CreateDataTwentyColumns(IWorksheet worksheet, int row)
        {
            string gCustData = $"{worksheet[row, 2].Value.Trim()},{worksheet[row, 3].Value.Trim()},";
            gCustData += $"{worksheet[row, 4].Value.Trim()},{worksheet[row, 6].Value.Trim()},{worksheet[row, 7].Value.Trim()},,,,,,,,,,,,,";
            gCustData += $"{worksheet[row, 8].Value.Trim()},{worksheet[row, 9].Value.Trim()},{worksheet[row, 10].Value.Trim()},{worksheet[row, 11].Value.Trim()},{worksheet[row, 12].Value.Trim()},,,,,,,,,,,,,";
            gCustData += $"{worksheet[row, 17].Value.Trim()},{worksheet[row, 18].Value.Trim()},{worksheet[row, 19].Value.Trim()},{worksheet[row, 13].Value.Trim()},{worksheet[row, 14].Value.Trim()},{worksheet[row, 15].Value.Trim()}";
            return gCustData;
        }
        /*The end is Write excel file has 20 column*/

        /*The Writing excel file has 24 column*/
        private void ProcessTwentyFourColumns(IWorksheet worksheet, int lastRow, string gPN, OracleConnection gConnectionScc)
        {
            for (int row = 6; row <= lastRow; row++)
            {
                if (CheckCellsNotEmpty(worksheet, row, new int[] { 1, 2, 3, 4, 5, 6, 19, 20, 21, 22 }))
                {
                    string gCustData = CreateDataTwentyFourColumns(worksheet, row);
                    ExecuteDatabase(worksheet, row, gPN, gCustData, gConnectionScc);
                }
            }
        }
        private string CreateDataTwentyFourColumns(IWorksheet worksheet, int row)
        {
            string gCustData = $"{worksheet[row, 2].Value.Trim()},{worksheet[row, 3].Value.Trim()},";
            gCustData += $"{worksheet[row, 4].Value.Trim()},{worksheet[row, 5].Value.Trim()},{worksheet[row, 6].Value.Trim()},,,,,,,,,,,,,";
            gCustData += $"{worksheet[row, 19].Value.Trim()},{worksheet[row, 20].Value.Trim()},{worksheet[row, 21].Value.Trim()},{worksheet[row, 22].Value.Trim()},,";
            return gCustData;
        }
        /*The end is Write excel file has 24 column*/

        /*The Writing excel file has 25 column*/
        private void ProcessTwentyFiveColumns(IWorksheet worksheet, int lastRow, string gPN, OracleConnection gConnectionScc)
        {
            for (int row = 6; row <= lastRow; row++)
            {
                if (CheckCellsNotEmpty(worksheet, row, new int[] { 1, 2, 3, 4, 5, 6, 19, 20, 21, 22, 23, 24, 25 }))
                {
                    string gCustData = CreateDataTwentyFiveColumns(worksheet, row);
                    ExecuteDatabase(worksheet, row, gPN, gCustData, gConnectionScc);
                }
            }
        }
        private string CreateDataTwentyFiveColumns(IWorksheet worksheet, int row)
        {
            string gCustData = $"{worksheet[row, 2].Value.Trim()},{worksheet[row, 3].Value.Trim()},";
            gCustData += $"{worksheet[row, 4].Value.Trim()},{worksheet[row, 5].Value.Trim()},{worksheet[row, 6].Value.Trim()},,,,,,,,,,,,,";
            gCustData += $"{worksheet[row, 19].Value.Trim()},{worksheet[row, 20].Value.Trim()},{worksheet[row, 21].Value.Trim()},{worksheet[row, 22].Value.Trim()},{worksheet[row, 23].Value.Trim()},{worksheet[row, 24].Value.Trim()},{worksheet[row, 25].Value.Trim()}";
            return gCustData;
        }
        /*The end is Write excel file has 25 column*/

        /*The Writing excel file has 41 column*/
        private void ProcessFortyOneColumns(IWorksheet worksheet, int lastRow, string gPN, OracleConnection gConnectionScc)
        {
            for (int row = 6; row <= lastRow; row++)
            {
                if (CheckCellsNotEmpty(worksheet, row, new int[] { 1, 2, 3, 4, 5, 6, 19, 20, 21, 22, 23, 36, 37, 38, 39, 40, 41, 42 }))
                {
                    string gCustData = CreateDataFortyOneColumns(worksheet, row);
                    ExecuteDatabase(worksheet, row, gPN, gCustData, gConnectionScc);
                }
            }
        }
        private string CreateDataFortyOneColumns(IWorksheet worksheet, int row)
        {
            string gCustData = $"{worksheet[row, 2].Value.Trim()},{worksheet[row, 3].Value.Trim()},";
            gCustData += $"{worksheet[row, 4].Value.Trim()},{worksheet[row, 5].Value.Trim()},{worksheet[row, 6].Value.Trim()},,,,,,,,,,,,,";
            gCustData += $"{worksheet[row, 19].Value.Trim()},{worksheet[row, 20].Value.Trim()},{worksheet[row, 21].Value.Trim()},{worksheet[row, 22].Value.Trim()},{worksheet[row, 23].Value.Trim()},,,,,,,,,,,,,{worksheet[row, 36].Value.Trim()},{worksheet[row, 37].Value.Trim()},{worksheet[row, 38].Value.Trim()},{worksheet[row, 39].Value.Trim()},{worksheet[row, 40].Value.Trim()},{worksheet[row, 41].Value.Trim()}";
            return gCustData;
        }
        /*The end is Write excel file has 41 column*/
        private bool CheckCellsNotEmpty(IWorksheet worksheet, int row, int[] columnIndices)
        {
            foreach (int col in columnIndices)
            {
                if (String.IsNullOrEmpty(worksheet[row, col].Value))
                {
                    return false;
                }
            }
            return true;
        }
        private void ExecuteDatabase(IWorksheet worksheet, int row, string gPN, string gCustData, OracleConnection gConnectionScc)
        {
            string gSQL = $"select cmac from SCC_CUST_MACSN_DATA where cmac = '{worksheet[row, 1].Value.Trim()}'";
            using (OracleCommand oraCommand = new OracleCommand(gSQL, gConnectionScc))
            {
                using (OracleDataReader dr = oraCommand.ExecuteReader())
                {
                    if (!dr.HasRows)
                    {
                        gSQL = $"Insert into SCC_CUST_MACSN_DATA(PART_NO, CUST_SN, CMAC, CUST_DATA, FLAG, CDT) VALUES('{gPN}', '{worksheet[row, 1].Value.Trim()}', '{worksheet[row, 1].Value.Trim()}', '{gCustData}', 'Y', sysdate)";
                    }
                    else
                    {
                        gSQL = $"UPDATE SCC_CUST_MACSN_DATA SET CUST_DATA = '{gCustData}', CDT = SYSDATE WHERE cmac = '{worksheet[row, 1].Value.Trim()}'";
                    }
                }
                using (OracleCommand updateCommand = new OracleCommand(gSQL, gConnectionScc))
                {
                    updateCommand.ExecuteNonQuery();
                }
                //Count += 1;
            }
        }
        public async Task WriteDataBosaNokia(string filePath)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {

                IApplication app = excelEngine.Excel;
                app.DefaultVersion = ExcelVersion.Excel2016;
                IWorkbook workbook = app.Workbooks.Open(filePath, ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];
                IRange usedRange = worksheet.UsedRange;
                int lastRow = usedRange.LastRow;

                var connections = OpenConnections();

                for (int row = 2; row <= lastRow; row++)
                {
                    if (CheckSerialNumber(worksheet[row, 2].Value))
                    {
                        var dr = await CheckPartNoBosa(connections.gConectionScc, worksheet[row, 1].Value);

                        while (dr.Read())
                        {
                            if (CheckSpecBosa(dr, worksheet, row))
                            {
                                if (!await CheckRecordExistsInNokia(connections.gConectionNokia, worksheet[row, 2].Value))
                                {
                                    await InsertIntoNokiaDatabase(connections.gConectionNokia, worksheet, row);
                                    await InsertIntoSfcsDatabase(connections.gConectionSFCS, worksheet, row);
                                }
                            }
                        }
                        dr.Close();
                    }
                }
                CloseConnections(connections);
            }
        }
        private (SqlConnection gConectionNokia, OracleConnection gConectionScc, OracleConnection gConectionSFCS) OpenConnections()
        {
            var gConectionNokia = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings["NOKIA_VN"].ConnectionString);
            gConectionNokia.Open();
            var gConectionScc = new OracleConnection(ConfigurationManager.ConnectionStrings["Gemtek_SCC"].ConnectionString);
            gConectionScc.Open();
            var gConectionSFCS = new OracleConnection(ConfigurationManager.ConnectionStrings["Gemtek_SFCS"].ConnectionString);
            gConectionSFCS.Open();

            return (gConectionNokia, gConectionScc, gConectionSFCS);
        }
        private void CloseConnections((SqlConnection gConectionNokia, OracleConnection gConectionScc, OracleConnection gConectionSFCS) connections)
        {
            connections.gConectionNokia.Close();
            connections.gConectionScc.Close();
            connections.gConectionSFCS.Close();
        }
        private async Task<OracleDataReader> CheckPartNoBosa(OracleConnection gConectionScc, string partNo)
        {
            string gSQL = "SELECT VENDER, BOSA_TYPE, VBRSLOPE1, VBRSLOPE2, AER, LOS, TD, TS, BOSAPN, P0, P1, OVBR, BOSA_DRIVER, APD_OFFSET, MOD_MAX, APC, MER, vbr, losd, bosa_id FROM SCC_BOSA_PN WHERE PART_NO = :partNo";
            using (OracleCommand oraCommand = new OracleCommand(gSQL, gConectionScc))
            {
                oraCommand.Parameters.Add(new OracleParameter("partNo", partNo));
                return (OracleDataReader)await oraCommand.ExecuteReaderAsync();
            }
        }
        private bool CheckSerialNumber(string serialNumber)
        {
            return serialNumber.Length > 5 && serialNumber.Length < 20;
        }
        private bool CheckSpecBosa(OracleDataReader dr, IWorksheet worksheet, int row)
        {
            return Convert.ToString(dr["VENDER"]).Equals(worksheet[row, 3].Value.Trim()) &&
                   Convert.ToString(dr["BOSA_TYPE"]).Equals(worksheet[row, 4].Value.Trim()) &&
                   Convert.ToString(dr["VBRSLOPE1"]).Equals(worksheet[row, 6].Value.Trim()) &&
                   Convert.ToString(dr["VBRSLOPE2"]).Equals(worksheet[row, 7].Value) &&
                   Convert.ToString(dr["AER"]).Equals(worksheet[row, 8].Value.Trim()) &&
                   Convert.ToString(dr["LOS"]).Equals(worksheet[row, 9].Value.Trim()) &&
                   Convert.ToString(dr["TD"]).Equals(worksheet[row, 10].Value.Trim()) &&
                   Convert.ToString(dr["TS"]).Equals(worksheet[row, 15].Value.Trim()) &&
                   Convert.ToString(dr["BOSAPN"]).Equals(worksheet[row, 17].Value.Trim()) &&
                   Convert.ToString(dr["P0"]).Equals(worksheet[row, 18].Value.Trim()) &&
                   Convert.ToString(dr["P1"]).Equals(worksheet[row, 19].Value.Trim()) &&
                   Convert.ToString(dr["BOSA_DRIVER"]).Equals(worksheet[row, 20].Value.Trim()) &&
                   Convert.ToString(dr["OVBR"]).Equals(worksheet[row, 21].Value.Trim()) &&
                   Convert.ToString(dr["APD_OFFSET"]).Equals(worksheet[row, 22].Value.Trim()) &&
                   Convert.ToString(dr["MOD_MAX"]).Equals(worksheet[row, 23].Value.Trim()) &&
                   Convert.ToString(dr["APC"]).Equals(worksheet[row, 24].Value.Trim()) &&
                   Convert.ToString(dr["MER"]).Equals(worksheet[row, 14].Value.Trim()) &&
                   Convert.ToString(dr["VBR"]).Equals(worksheet[row, 5].Value.Trim()) &&
                   (Convert.ToString(dr["losd"]) == "" ? "NULL" : Convert.ToString(dr["losd"])).Equals(worksheet[row, 25].Value.Trim()); 
                   //(Convert.ToString(dr["bosa_id"]) == "" ? "NULL" : Convert.ToString(dr["bosa_id"])).Equals(worksheet[row, 26].Value.Trim());
        }
        private async Task<bool> CheckRecordExistsInNokia(SqlConnection gConectionNokia, string serialNumber)
        {
            string gSQL = "SELECT Serialnumber FROM [ONT1.4.1].[dbo].[BOSAInfo_Table] WHERE Serialnumber = @serialNumber";
            SqlCommand command = new SqlCommand(gSQL, gConectionNokia);
            command.Parameters.Add(new SqlParameter("@serialNumber", serialNumber));
            SqlDataReader reader = await command.ExecuteReaderAsync();
            bool exists = await reader.ReadAsync();
            reader.Close();
            return exists;
        }
        private async Task InsertIntoNokiaDatabase(SqlConnection gConectionNokia, IWorksheet worksheet, int row)
        {
            string gSQL = "" + "insert into [ONT1.4.1].[dbo].[BOSAInfo_Table] (Serialnumber,Vendor,BOSAType,Vbr,VbrSlope1,VbrSlope2,Aer,Los,";
            gSQL = gSQL + "Td,InputDate,VBRretreat,BOSACode,Mer,ts,Ith,BOSAPN,P0,P1,BOSA_Driver,Ovbr,APD_offset,Mod_MAX,APC,CDT,LD_Driver,BOSAID)";
            gSQL = gSQL + "Values('" + worksheet[row, 2].Value + "','" + worksheet[row, 3].Value + "','" + worksheet[row, 4].Value + "'";
            gSQL = gSQL + ",'" + worksheet[row, 5].Value + "'," + worksheet[row, 6].Value + "," + worksheet[row, 7].Value + "";
            gSQL = gSQL + "," + worksheet[row, 8].Value + "," + worksheet[row, 9].Value + "," + worksheet[row, 10].Value + ",";
            gSQL = gSQL + "getdate()," + worksheet[row, 12].Value + ",'" + worksheet[row, 13].Value + "',";
            gSQL = gSQL + "'" + worksheet[row, 14].Value + "','" + worksheet[row, 15].Value + "'," + worksheet[row, 16].Value + ",";
            gSQL = gSQL + "'" + worksheet[row, 17].Value + "'," + worksheet[row, 18].Value + "," + worksheet[row, 19].Value + ",";
            gSQL = gSQL + "'" + worksheet[row, 20].Value + "'," + worksheet[row, 21].Value + ",'" + worksheet[row, 22].Value + "',";
            gSQL = gSQL + "'" + worksheet[row, 23].Value + "'," + worksheet[row, 24].Value + ",getdate(),'" + worksheet[row, 25].Value + "','" + worksheet[row, 26].Value + "')";
            SqlCommand command = new SqlCommand(gSQL, gConectionNokia);
            await command.ExecuteNonQueryAsync();
        }
        private async Task InsertIntoSfcsDatabase(OracleConnection gConectionSFCS, IWorksheet worksheet, int row)
        {
            string gSQL = "INSERT INTO BOSAInfo_Table (Serialnumber, Vendor, BOSAType, Vbr, VbrSlope1, VbrSlope2, Aer, Los, Td, InputDate, VBRretreat, BOSACode, Mer, ts, Ith, BOSAPN, P0, P1, BOSA_Driver, Ovbr, APD_offset, Mod_MAX, APC, CDT) " +
                          "VALUES (:Serialnumber, :Vendor, :BOSAType, :Vbr, :VbrSlope1, :VbrSlope2, :Aer, :Los, :Td, SYSDATE, :VBRretreat, :BOSACode, :Mer, :ts, :Ith, :BOSAPN, :P0, :P1, :BOSA_Driver, :Ovbr, :APD_offset, :Mod_MAX, :APC, SYSDATE)";
            OracleCommand command = new OracleCommand(gSQL, gConectionSFCS);
            command.Parameters.Add(new OracleParameter("Serialnumber", worksheet[row, 2].Value));
            command.Parameters.Add(new OracleParameter("Vendor", worksheet[row, 3].Value));
            command.Parameters.Add(new OracleParameter("BOSAType", worksheet[row, 4].Value));
            command.Parameters.Add(new OracleParameter("Vbr", worksheet[row, 5].Value));
            command.Parameters.Add(new OracleParameter("VbrSlope1", worksheet[row, 6].Value));
            command.Parameters.Add(new OracleParameter("VbrSlope2", worksheet[row, 7].Value));
            command.Parameters.Add(new OracleParameter("Aer", worksheet[row, 8].Value));
            command.Parameters.Add(new OracleParameter("Los", worksheet[row, 9].Value));
            command.Parameters.Add(new OracleParameter("Td", worksheet[row, 10].Value));
            command.Parameters.Add(new OracleParameter("VBRretreat", worksheet[row, 12].Value));
            command.Parameters.Add(new OracleParameter("BOSACode", worksheet[row, 13].Value));
            command.Parameters.Add(new OracleParameter("Mer", worksheet[row, 14].Value));
            command.Parameters.Add(new OracleParameter("ts", worksheet[row, 15].Value));
            command.Parameters.Add(new OracleParameter("Ith", worksheet[row, 16].Value));
            command.Parameters.Add(new OracleParameter("BOSAPN", worksheet[row, 17].Value));
            command.Parameters.Add(new OracleParameter("P0", worksheet[row, 18].Value));
            command.Parameters.Add(new OracleParameter("P1", worksheet[row, 19].Value));
            command.Parameters.Add(new OracleParameter("BOSA_Driver", worksheet[row, 20].Value));
            command.Parameters.Add(new OracleParameter("Ovbr", worksheet[row, 21].Value));
            command.Parameters.Add(new OracleParameter("APD_offset", worksheet[row, 22].Value));
            command.Parameters.Add(new OracleParameter("Mod_MAX", worksheet[row, 23].Value));
            command.Parameters.Add(new OracleParameter("APC", worksheet[row, 24].Value));
            await command.ExecuteNonQueryAsync();
        }
        public async Task<List<string>> GetSheetName(string filePath)
        {
            List<string> sheetNames = new List<string>();

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication app = excelEngine.Excel;
                app.DefaultVersion = ExcelVersion.Excel2016;
                IWorkbook workbook = app.Workbooks.Open(filePath, ExcelOpenType.Automatic);
                foreach (IWorksheet sheet in workbook.Worksheets)
                {
                    sheetNames.Add(sheet.Name);
                }
            }
            return sheetNames;
        }
        public async Task<int> GetTotalColumns(string filePath)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication app = excelEngine.Excel;
                app.DefaultVersion = ExcelVersion.Excel2016;
                IWorkbook workbook = app.Workbooks.Open(filePath, ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];
                IRange usedRange = worksheet.UsedRange;
                int NumberColumn = 0;
                for (int row = 1; row <= usedRange.LastRow; row++)
                {
                    for (int col = 1; col <= usedRange.LastColumn; col++)
                    {
                        if (worksheet[row, col].Value != null && !string.IsNullOrEmpty(worksheet[row, col].Value.ToString().Trim()))
                        {
                            if (col > NumberColumn)
                            {
                                NumberColumn = col;
                            }
                        }
                    }
                }
                return NumberColumn;
            }
        }
    }
}
