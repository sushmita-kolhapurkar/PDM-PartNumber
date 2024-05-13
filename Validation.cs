using System;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using EPDM.Interop.epdm;

namespace AddIn_PartNumber
{
    class Validation
    {
        public bool searchDB(object oSearchTerm)
        {
            bool isValid = true;

            //string sqlConnString = "Server=SV-SQL01;Database=LUCE;User Id=INTERACTIVE; Password=LUCE1234$;Encrypt=false";
            string sqlConnString = "Server=SV-SQL01\\DEV;Database=LUCEDEV;User Id=INTERACTIVE; Password=LUCE1234$";
          //  string sqlConnString = ConfigurationManager.ConnectionStrings["LuceConnection"].ConnectionString;
            SqlConnection conn = new SqlConnection(sqlConnString);
            conn.Open();
            SqlCommand cmd;

            string sqlQuery =
            "IF EXISTS(SELECT PART.id, DESCRIPTION FROM PART " +
            "WHERE " +
            "(PART.ID  = '" + oSearchTerm + "' or PART.ID  like '%[-| ]" + oSearchTerm + "' or PART.ID  like '" + oSearchTerm + "[-| ]%') " +
            "and (PART.ID not like '3G%' and PART.ID not like 'fha%' and PART.ID not like 'fta%' and PART.ID not like 'fga%' " +
            "and PART.ID not like 'fia%' and PART.ID not like 'fgr%' and PART.ID not like 'fxa%') " +
            "or (DESCRIPTION = '" + oSearchTerm + "' or DESCRIPTION like '%[-| ]" + oSearchTerm + "[-| ]%') " +
            ") SELECT 0 ELSE SELECT 1";
            cmd = new SqlCommand(sqlQuery, conn);
            object ret = cmd.ExecuteScalar();
            isValid = Convert.ToBoolean(ret);

            conn.Close();

            return isValid;
        }
        public bool searchPDM_FileName(IEdmVault5 vault5, object oSearchTerm)
        {
            bool isValid = true;
            IEdmVault21 vault = (IEdmVault21)vault5;
            IEdmSearch9 search9 = (IEdmSearch9)vault.CreateSearch2();
            IEdmSearchResult5 result;

                search9.Clear();
                string searchTerm = "" + oSearchTerm + ".* OR " + oSearchTerm + "- OR -" + oSearchTerm + "";
                search9.FileName = searchTerm;
                search9.FindFolders = false;
                search9.FindFolders = true;

                result = search9.GetFirstResult();

            if (result != null)
            {
                var regex = new Regex(@"\b[^\d\r\n]?" + oSearchTerm + "\\b");
                bool isMatch = regex.Match(result.Name).Success;

                isValid = !isMatch;
            }
            return isValid;
        }
        public bool searchPDM_Variable(IEdmVault5 vault5, object oSearchTerm)
        {
            bool isValid = true;
            IEdmVault21 vault = (IEdmVault21)vault5;
            IEdmSearch5 search5 = (IEdmSearch5)vault.CreateUtility(EdmUtility.EdmUtil_Search);
            IEdmSearch9 search9 = (IEdmSearch9)vault.CreateSearch2();

            IEdmSearchResult5 result;

                search9.Clear();
                string searchTerm = "=" + oSearchTerm + " OR *" + oSearchTerm + "- OR -" + oSearchTerm + "";

                string[] VarNames = { "PartNumber", "Number" };
                search9.AddMultiVariableCondition(VarNames, searchTerm);
                result = search9.GetFirstResult();

                if (result != null)
                {
                isValid = false;
            }
            return isValid;
        }
    }
}
