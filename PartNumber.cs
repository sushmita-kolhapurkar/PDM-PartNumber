using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using EPDM.Interop.epdm;
using EPDM.Interop.EPDMResultCode;

namespace AddIn_PartNumber
{
    public class PartNumber : IEdmAddIn5
    {
        public string FileName = "C:\\Users\\Sushmita\\OneDrive - 3GLighting\\Desktop\\PartNos.xls"; //change this

        void IEdmAddIn5.GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            try
            {
                poInfo.mbsAddInName = "PartNumber";
                poInfo.mbsCompany = "3G Lighting";
                poInfo.mbsDescription = "Auto-generate part numbers for new designs - Manufactured, Purchased or R&D";
                poInfo.mlAddInVersion = 1;
                poInfo.mlRequiredVersionMajor = 5;
                poInfo.mlRequiredVersionMinor = 2;

                poCmdMgr.AddHook(EdmCmdType.EdmCmd_CardInput);
                poCmdMgr.AddHook(EdmCmdType.EdmCmd_CardButton);
              //  poCmdMgr.AddHook(EdmCmdType.EdmCmd_PostUnlock);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        void IEdmAddIn5.OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            IEdmVault5 vault = (IEdmVault5)poCmd.mpoVault;
            // Microsoft.CSharp.CompilerServices.StaticLocalInitFlag static_OnCmd_VariableChangeInProgress_Init = new Microsoft.VisualBasic.CompilerServices.StaticLocalInitFlag();
            // https://help.solidworks.com/2020/english/api/epdmapi/Change_Card_Variables_Addin_Example_CSharp.htm

            switch (poCmd.meCmdType)
            {
                /* For Dropdown change event */
                case EdmCmdType.EdmCmd_CardInput:
                    {
                        try
                        {
                            #region Var_inProg
                            /*bool static_OnCmd_VariableChangeInProgress;

                            lock (static_OnCmd_VariableChangeInProgress_Init)
                            {
                                try
                                {
                                    if (InitStaticVariableHelper(static_OnCmd_VariableChangeInProgress_Init))
                                    {
                                        static_OnCmd_VariableChangeInProgress = false;
                                    }
                                }
                                finally
                                {
                                    static_OnCmd_VariableChangeInProgress_Init.State = 1;
                                }
                            } */
                            #endregion

                            if (poCmd.mbsComment == "Part Stage")
                            {
                                IEdmEnumeratorVariable5 vars = (IEdmEnumeratorVariable5)poCmd.mpoExtra;
                                string Config = ((EdmCmdData)ppoData.GetValue(0)).mbsStrData1;

                                vars.GetVar("Part Stage", Config, out object oProdType);

                                if (oProdType.ToString() == "Manufactured")
                                {
                                    int iPartNo = GetManufacturedPNo(FileName);
                                    vars.SetVar("Number", Config, iPartNo.ToString(), true);
                                }
                                else if (oProdType.ToString() == "Purchased")
                                {
                                    SetPurchasedPNo(Config, vars);
                                }
                                else if (oProdType.ToString() == "R&D")
                                {
                                    vars.SetVar("Number", Config, "", true);
                                }
                            }
                            else
                            if (poCmd.mbsComment == "Category")
                            {
                                IEdmEnumeratorVariable5 vars = (IEdmEnumeratorVariable5)poCmd.mpoExtra;
                                string Config = ((EdmCmdData)ppoData.GetValue(0)).mbsStrData1;

                                vars.GetVar("Part Stage", Config, out object oProdType);
                                if (oProdType.ToString() == "Purchased")
                                {
                                    vars.GetVar("Category", Config, out object oCategory);
                                    string strNewPartId = GetPurchasedPno(oCategory);

                                    //Check if it was a 'Manufactured' part before
                                    vars.GetVar("Number", Config, out object oPartNo);
                                    vars.GetVar("Description", Config, out object oDesc);

                                    string strDesc = GetDesc(oPartNo, oDesc);
                                    if (strDesc != "")
                                        vars.SetVar("Description", Config, strDesc, true);
                                    //---//

                                    vars.SetVar("Number", Config, strNewPartId, true);
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        break;
                    }

                case EdmCmdType.EdmCmd_CardButton:
                    {
                        if (poCmd.mbsComment == "AddIn_OnSave")
                        {
                            Validation val = new Validation();
                            bool isValid = true;

                            IEdmEnumeratorVariable8 vars = (IEdmEnumeratorVariable8)poCmd.mpoExtra;
                            string Config = ((EdmCmdData)ppoData.GetValue(0)).mbsStrData1;

                            vars.GetVar("Part Stage", Config, out object oProdType);
                            vars.GetVar("Number", Config, out object oPartNo);
                            int.TryParse(oPartNo.ToString(), out int iNo);

                            //Validation
                            if (oPartNo != null)
                            {
                                isValid = val.searchDB(oPartNo);
                                if (isValid)
                                    isValid = val.searchPDM_FileName(vault, oPartNo);
                                if (isValid)
                                    isValid = val.searchPDM_Variable(vault, oPartNo);

                                if (!isValid)
                                {
                                    MessageBox.Show("This Part Id already exists");
                                    if (oProdType.ToString() == "Manufactured")
                                    {
                                        DelManufacturedPNo(FileName, oPartNo.ToString());
                                        int iPartNo = GetManufacturedPNo(FileName);
                                        vars.SetVar("Number", Config, iPartNo.ToString(), true);
                                        DelManufacturedPNo(FileName, oPartNo.ToString());
                                    }
                                    else if (oProdType.ToString() == "Purchased")
                                    {
                                        SetPurchasedPNo(Config, vars);
                                    }
                                    vars.CloseFile(isValid);
                                    EdmCmdData cmdData = (EdmCmdData)ppoData.GetValue(0);
                                    cmdData.mlLongData1 = (int)EdmCardFlag.EdmCF_CloseDlgCancel;
                                    //break;
                                }
                                else
                                {
                                    string[] fields = { "Number" };

                                    //Validate fields
                                    isValid = CheckFields(vars, Config, fields);

                                    if (isValid)
                                    {
                                        //Delete from excel
                                        DelManufacturedPNo(FileName, oPartNo.ToString());
                                        vars.CloseFile(isValid);
                                        if (isValid)
                                        {
                                            EdmCmdData cmdData = (EdmCmdData)ppoData.GetValue(0);
                                            cmdData.mlLongData1 = (int)EdmCardFlag.EdmCF_CloseDlgOK;
                                            ppoData.SetValue(cmdData, 0);
                                            MessageBox.Show("The file has been saved");
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }

                                }
                            }
                        }
                        break;
                    }
                #region CheckIn Validation
                //case EdmCmdType.EdmCmd_PostUnlock:
                //    {
                //        try
                //        {
                //            string FilePath = ((EdmCmdData)ppoData.GetValue(0)).mbsStrData1;

                //            if (FilePath.ToLower().EndsWith(".sldprt"))
                //            {
                //                IEdmFile6 file = (IEdmFile6)vault.GetFileFromPath(FilePath, out IEdmFolder5 folder);

                //                if (!file.IsLocked)
                //                {
                //                    file.LockFile(folder.ID, poCmd.mlParentWnd, ((int)EdmLockFlag.EdmLock_Simple));
                //                }

                //                IEdmEnumeratorVariable8 vars = (IEdmEnumeratorVariable8)file.GetEnumeratorVariable();

                //                EdmStrLst5 strConfig = file.GetConfigurations();
                //                IEdmPos5 pos_conf = strConfig.GetHeadPosition();
                //                EdmCmdData cmdData = (EdmCmdData)ppoData.GetValue(0);

                //                while (!pos_conf.IsNull)
                //                {
                //                    string Config = strConfig.GetNext(pos_conf);

                //                    vars.GetVar("Number", Config, out object oPartNo);

                //                    Validation

                //                    Validation val = new Validation();
                //                    bool isValid = true;
                //                    isValid = val.searchDB(oPartNo);
                //                    if (isValid)
                //                        isValid = val.searchPDM_FileName(vault, oPartNo);
                //                    if (isValid)
                //                        isValid = val.searchPDM_Variable(vault, oPartNo);

                //                    if (!isValid)
                //                    {
                //                        MessageBox.Show("This Part Id already exists");
                //                    }
                //                    vars.CloseFile(isValid);
                //                }
                //                cmdData.mlLongData1 = (int)EdmCardFlag.EdmCF_Nothing;
                //            }
                //        }
                //    catch (Exception e)
                //    {
                //    }
                //    break;
                //}
                #endregion

                default:
                    break;
            }
        }
        public static void SetPurchasedPNo(string Config, IEdmEnumeratorVariable5 vars)
        {
            //Check if Category has been filled 
            vars.GetVar("Category", Config, out object oCategory);
            if (oCategory == null || oCategory.ToString() == "")
            {
                MessageBox.Show("Select 'Category'");
                vars.SetVar("Number", Config, "", true);
            }
            else
            {
                string strNewPartId = GetPurchasedPno(oCategory);

                //Check if it was a 'Fabricated' part before
                vars.GetVar("Number", Config, out object oPartNo);
                vars.GetVar("Description", Config, out object oDesc);

                string strDesc = GetDesc(oPartNo, oDesc);
                if (strDesc != "")
                    vars.SetVar("Description", Config, strDesc, true);
                //---//

                vars.SetVar("Number", Config, strNewPartId, true);
            }
        }

        public static int GetManufacturedPNo(string fileName)
        {
            int iRet = 0;

            try
            {
                string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"" + fileName + "\";Extended Properties=\"Excel 12.0;IMEX=3;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\"";

                using (var conn = new OleDbConnection(connString))
                {
                    conn.Open();

                    var sheets = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    if (sheets != null)
                    {
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = "SELECT * FROM [Sheet1$]";

                            var adapter = new OleDbDataAdapter(cmd);
                            var ds = new DataSet();
                            adapter.Fill(ds);

                            DataTable dtExcel = ds.Tables[0];

                            for (int iCol = 0; iCol < dtExcel.Columns.Count; iCol++)
                            {
                                for (int iRow = 0; iRow < dtExcel.Rows.Count; iRow++)
                                {
                                    string sValue = dtExcel.Rows[iRow][iCol].ToString();
                                    if (sValue != "")
                                    {
                                        int.TryParse(sValue, out iRet);
                                        if (iRet != 0)
                                            goto RET;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
            }

        RET:
            return iRet;
        }
        public static int DelManufacturedPNo(string fileName, string sPartNo)
        {
            int iRet = 0;
            //FileStream stream = null;

            try
            {
                //Check and get only part number from the textbox value
                System.Text.RegularExpressions.Match mtc = Regex.Match(sPartNo, "(?<!\\d)\\d{5}(?!\\d)");
                if (mtc.Length > 0)
                {
                    int.TryParse(mtc.Groups[0].Value, out int iNo);
                    string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"" + fileName + "\";Extended Properties=\"Excel 12.0;IMEX=3;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\"";

                    using (var conn = new OleDbConnection(connString))
                    {
                        conn.Open();

                        var sheets = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                        if (sheets != null)
                        {
                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.CommandText = "SELECT * FROM [Sheet1$]";

                                var adapter = new OleDbDataAdapter(cmd);
                                var ds = new DataSet();
                                adapter.Fill(ds);

                                DataTable dtExcel = ds.Tables[0];

                                // var matchingRows = dtExcel.AsEnumerable().Where(row => row.ItemArray.Contains(iNo)).ToList();

                                for (int iCol = 0; iCol < dtExcel.Columns.Count; iCol++)
                                {
                                    string query = String.Format("UPDATE [Sheet1$] SET F" + (iCol + 1) + "=0 WHERE F" + (iCol + 1) + "=" + iNo);
                                    OleDbCommand cmds = new OleDbCommand(query, conn);
                                    int isSuccess = cmds.ExecuteNonQuery();

                                    if (isSuccess != 0)
                                        break;
                                }
                            }
                        }
                        conn.Close();
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
            }

            return iRet;
        }
        public static string GetPurchasedPno(object oCategory)
        {
            string strCategory = "", strPartId = "", strNewPartId = "";
            int iPartId = 0, iNewPartId = 0;
            string sqlConnString = "Server=SV-SQL01\\DEV;Database=LUCEDEV;User Id=INTERACTIVE; Password=LUCE1234$;"; //change this
            SqlConnection conn = new SqlConnection(sqlConnString);
            conn.Open();

            string sqlQuery = "SELECT top 1 id from LUCEDEV.dbo.PART where id like '" + oCategory.ToString() + "[0-9]%' order by id desc;";
            SqlCommand cmd = new SqlCommand(sqlQuery, conn);

            object ret = cmd.ExecuteScalar();
            if (ret != null)
            {
                string strPId_fromDb = ret.ToString();

                if (strPId_fromDb.Contains('-'))
                {
                    strPId_fromDb = strPId_fromDb.Substring(0, strPId_fromDb.IndexOf('-'));
                }

                strCategory = Regex.Match(strPId_fromDb, @"\D+").Value;
                strPartId = Regex.Match(strPId_fromDb, @"\d+").Value;
                int.TryParse(strPartId, out iPartId);

                int iLen = strPartId.Length;

                iNewPartId = iPartId + 1;
                strPartId = iNewPartId.ToString();

                strNewPartId = strCategory + strPartId.PadLeft(iLen, '0');
            }
            else
            {
                MessageBox.Show("No relevant part ID found in the database");
            }

            return strNewPartId;
        }
        public static string GetDesc(object oPartNo, object oDesc)
        {
            string strDesc = "";
            Regex pattern = new Regex("(?<!\\d)\\d{5}(?!\\d)");
            System.Text.RegularExpressions.Match mtc = pattern.Match(oPartNo.ToString());
            if (mtc.Length > 0)
            {
                System.Text.RegularExpressions.Match mtc2 = pattern.Match(oDesc.ToString());

                strDesc = oDesc.ToString();

                if (mtc2.Length > 0)
                {
                    if (mtc2.Groups[0].Value == oPartNo.ToString())
                    {
                        strDesc = strDesc.Replace("-" + oPartNo.ToString(), "");
                    }
                }
                strDesc += "-" + oPartNo.ToString();
            }

            return strDesc;
        }

        public static bool CheckFields(IEdmEnumeratorVariable5 vars, string Config, string[] fields)
        {
            bool isNull = false;
            List<KeyValuePair<string, string>> kvFields = new List<KeyValuePair<string, string>>();

            foreach (string fieldname in fields)
            {
                vars.GetVar(fieldname, Config, out object Ret);
                kvFields.Add(new KeyValuePair<string, string>(fieldname, Ret.ToString()));
            }

            string strNullFields = "";

            //Validation check
            foreach (KeyValuePair<string, string> kv in kvFields)
            {
                if (kv.Value.ToString() == "" || kv.Value == null)
                {
                    isNull = true;
                    strNullFields += kv.Key;
                }
            }

            if (isNull)
                MessageBox.Show("Enter the following fields: " + strNullFields);

            return !isNull;
        }
    }
}
