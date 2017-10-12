using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

//see page 51
//using System.Data.SqlClient;
using System.Data.OleDb;


/// <summary>
/// Summary description for MyAdoHelper
/// פעולות עזר לשימוש במסד נתונים  מסוג 
/// SQL SERVER
///  App_Data המסד ממוקם בתקיה 
/// </summary>
public class MyAdoHelper
{
    public MyAdoHelper()
    {
        //
        // TODO: Add constructor logic here
        //
    }



    //public static SqlConnection ConnectToDb(string fileName)
    public static OleDbConnection ConnectToDb(string fileName)
    {
        // https://docs.microsoft.com/en-us/sql/ado/guide/appendixes/microsoft-ole-db-provider-for-sql-server
        //"Provider=SQLOLEDB;Data Source=serverName;"
        //Initial Catalog = databaseName;
        //  User ID = MyUserID; Password = MyPassword; "

        //https://www.connectionstrings.com/microsoft-ole-db-provider-for-sql-server-sqloledb/
        //Provider = sqloledb; Data Source = myServerAddress; Initial Catalog = myDataBase;
        //User Id = myUsername; Password = myPassword;

        //https://www.connectionstrings.com/access-2013/
        //Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\myFolder\myAccessFile.accdb;
        //Persist Security Info = False;


        string path = HttpContext.Current.Server.MapPath("App_Data/");//מיקום מסד בפורוייקט
        path += fileName;

        //string path = HttpContext.Current.Server.MapPath("App_Data/" + fileName);//מאתר את מיקום מסד הנתונים מהשורש ועד התקייה בה ממוקם המסד
        //orig connectio string
        //string connString = @"Data Source=.\SQLEXPRESS;AttachDbFilename=" +
        //path +
        //";Integrated Security=True;User Instance=True";

        //modify connection string
        string connString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" +
        path +
        ";Persist Security Info = False;";

        //SqlConnection conn = new SqlConnection(connString);
        OleDbConnection conn = new OleDbConnection(connString);
        
        return conn;
    }
    
    
    
    /// <summary>
    /// To Execute update / insert / delete queries
    ///  הפעולה מקבלת שם קובץ ומשפט לביצוע ומבצעת את הפעולה על המסד
    /// </summary>
    public static void DoQuery(string fileName, string sql)//הפעולה מקבלת שם מסד נתונים ומחרוזת מחיקה/ הוספה/ עדכון
    //ומבצעת את הפקודה על המסד הפיזי
    {
        //SqlConnection conn = ConnectToDb(fileName);
        OleDbConnection conn = ConnectToDb(fileName);
        try
        {
            conn.Open();
        }
        catch (OleDbException e)
        {
            string errorMessages = "";

            for (int i = 0; i < e.Errors.Count; i++)
            {
                errorMessages += "Index #" + i + "\n" +
                                 "Message: " + e.Errors[i].Message + "\n" +
                                 "NativeError: " + e.Errors[i].NativeError + "\n" +
                                 "Source: " + e.Errors[i].Source + "\n" +
                                 "SQLState: " + e.Errors[i].SQLState + "\n";
            }
            Console.WriteLine("DoQuery::conn.Open() An exception occurred. Please contact your system administrator." + errorMessages);
        }

        //SqlCommand com = new SqlCommand(sql, conn);
        OleDbCommand com = new OleDbCommand(sql, conn);

        try
        {
            com.ExecuteNonQuery();
        }
        catch (OleDbException e)
        {
            string errorMessages = "";

            for (int i = 0; i < e.Errors.Count; i++)
            {
                errorMessages += "Index #" + i + "\n" +
                                 "Message: " + e.Errors[i].Message + "\n" +
                                 "NativeError: " + e.Errors[i].NativeError + "\n" +
                                 "Source: " + e.Errors[i].Source + "\n" +
                                 "SQLState: " + e.Errors[i].SQLState + "\n";
            }
            Console.WriteLine("DoQuery::com.ExecuteNonQuery() An exception occurred. Please contact your system administrator." + errorMessages);
        }


        com.Dispose();
        conn.Close();
    }
    
    
    
    /// <summary>
    /// To Execute update / insert / delete queries
    ///  הפעולה מקבלת שם קובץ ומשפט לביצוע ומחזירה את מספר השורות שהושפעו מביצוע הפעולה
    /// </summary>
    public static int RowsAffected(string fileName, string sql)//הפעולה מקבלת מסלול מסד נתונים ופקודת עדכון
    //ומבצעת את הפקודה על המסד הפיזי
    {
        //SqlConnection conn = ConnectToDb(fileName);
        OleDbConnection conn = ConnectToDb(fileName);

        conn.Open();
        //SqlCommand com = new SqlCommand(sql, conn);
        OleDbCommand com = new OleDbCommand(sql, conn);

        int rowsA = com.ExecuteNonQuery();
        conn.Close();
        return rowsA;
    }
    
    
    /// <summary>
    /// הפעולה מקבלת שם קובץ ומשפט לחיפוש ערך - מחזירה אמת אם הערך נמצא ושקר אחרת
    /// </summary>
    public static bool IsExist(string fileName, string sql)//הפעולה מקבלת שם קובץ ומשפט בחירת נתון ומחזירה אמת אם הנתונים קיימים ושקר אחרת
    {
        //SqlConnection conn = ConnectToDb(fileName);
        OleDbConnection conn;
        conn = ConnectToDb(fileName);

        ;
        //https://msdn.microsoft.com/en-us/library/system.data.oledb.oledbexception(v=vs.90).aspx


        //SqlCommand com = new SqlCommand(sql, conn);
        OleDbCommand com = new OleDbCommand(sql, conn);

        try
        {
            conn.Open();
        }
        catch (OleDbException e)
        {
            string errorMessages = "";

            for (int i = 0; i < e.Errors.Count; i++)
            {
                errorMessages += "Index #" + i + "\n" +
                                 "Message: " + e.Errors[i].Message + "\n" +
                                 "NativeError: " + e.Errors[i].NativeError + "\n" +
                                 "Source: " + e.Errors[i].Source + "\n" +
                                 "SQLState: " + e.Errors[i].SQLState + "\n";
            }
            Console.WriteLine("IsExist:: conn.Open()An exception occurred. Please contact your system administrator."+ errorMessages);
        }


        
        bool b_flag = true;
        OleDbDataReader m_data;
        bool found = false;
        //SqlDataReader data = com.ExecuteReader();
        try
        {
            

            m_data = com.ExecuteReader();

            if (b_flag == true)
            {
                //undersatand what to do here?
                found = m_data.Read();// אם יש נתונים לקריאה יושם אמת אחרת שקר - הערך קיים במסד הנתונים
            }


        }
        catch (OleDbException e)
        {

            b_flag = false;

            string errorMessages = "";

            for (int i = 0; i < e.Errors.Count; i++)
            {
                errorMessages += "Index #" + i + "\n" +
                                 "Message: " + e.Errors[i].Message + "\n" +
                                 "NativeError: " + e.Errors[i].NativeError + "\n" +
                                 "Source: " + e.Errors[i].Source + "\n" +
                                 "SQLState: " + e.Errors[i].SQLState + "\n";
            }
            Console.WriteLine("IsExist:: com.ExecuteReader An exception occurred. Please contact your system administrator."+ errorMessages);
        }


       

        

        conn.Close();
        return found;
    }



    public static DataTable ExecuteDataTable(string fileName, string sql)
    {
        //SqlConnection conn = ConnectToDb(fileName);
        OleDbConnection conn = ConnectToDb(fileName);
        conn.Open();
        //SqlDataAdapter tableAdapter = new SqlDataAdapter(sql, conn);
        OleDbDataAdapter tableAdapter = new OleDbDataAdapter(sql, conn);

        DataTable dt = new DataTable();
        tableAdapter.Fill(dt);
        return dt;
    }



    public static void ExecuteNonQuery(string fileName, string sql)
    {
        //SqlConnection conn = ConnectToDb(fileName);
        OleDbConnection conn = ConnectToDb(fileName);

        conn.Open();
        //SqlCommand command = new SqlCommand(sql, conn);
        OleDbCommand cmd = new OleDbCommand(sql, conn);
        cmd.ExecuteNonQuery();
        conn.Close();
    }



    public static string printDataTable(string fileName, string sql)//הפעולה מקבלת שם קובץ ומשפט בחירת נתון ומחזירה אמת אם הנתונים קיימים ושקר אחרת
    {
        DataTable dt = ExecuteDataTable(fileName, sql);
        string printStr = "<table border='1'>";
        foreach (DataRow row in dt.Rows)
        {
            printStr += "<tr>";
            foreach (object myItemArray in row.ItemArray)
            {
                printStr += "<td>" + myItemArray.ToString() + "</td>";
            }
            printStr += "</tr>";
        }
        printStr += "</table>";
        return printStr;
    }
}