using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using Microsoft.Win32;
using NAWXDBCINFOIOLib;
using System.Web.Configuration;

/// <summary>
/// wTools 的摘要描述
/// </summary>
public class wTools
{
    //private ServiceReference1.ServiceSoapClient ws = new ServiceReference1.ServiceSoapClient();
    private string subkey;
    private string project;
    private string bpmconn;
    private string jbconn;
    private int dolog;
    private string strjbcmd = "";
    public wTools()
    {
        //
        // TODO: 在此加入建構函式的程式碼
        //
        subkey = WebConfigurationManager.ConnectionStrings["NTRegSubKey"].ConnectionString;
        project = WebConfigurationManager.ConnectionStrings["NTRegProject"].ConnectionString;
        bpmconn = WebConfigurationManager.ConnectionStrings["NTBPMconn"].ConnectionString;
        jbconn = WebConfigurationManager.ConnectionStrings["NTJBconn"].ConnectionString;
        dolog = Convert.ToInt32(WebConfigurationManager.ConnectionStrings["dolog"].ConnectionString);
    }

    /// <summary>
    /// 計算間隔秒數 或天數
    /// </summary>
    /// <param name="Sdate">開始時間</param>
    /// <param name="Edate">結束時間</param>
    /// <param name="Flag">傳回的類別 EX: sec , day,min,hour</param>
    /// <returns>Result</returns>
    public long Get_DateDiff(DateTime Sdate, DateTime Edate, string Flag)
    {
        TimeSpan t = Edate - Sdate;
        long Result = 0;

        switch (Flag)
        {
            case "sec"://傳回間隔的秒數
                Result = (long)t.TotalSeconds;
                break;
            case "day"://傳回間隔的天數
                Result = (long)t.TotalDays;
                break;
            case "min": //傳回間隔的分鐘數
                Result = (long)t.TotalMinutes;
                break;
            case "hour": //傳回間隔的分鐘數
                Result = (long)t.TotalHours;
                break;
        }
        return Result;
    }

    /// <summary>
    /// 計算間隔秒數 或天數
    /// </summary>
    /// <param name="duration">間隔秒數 或天數</param>
    /// <returns>HHMMSS</returns>
    public string ConvertDurationToHHMMSS(long duration)
    {
        //轉換成秒
        int SS = (int)duration % 60;
        //轉換成分鐘
        int MM = (int)(duration / 60) % 60;
        //轉換成時間
        long HH = (long)(duration / 60 / 60);
        //使用String Format 令到他們 會有 "0"/"00" 為開始
        return String.Format("{0:#00}:{1:#00}:{2:#00}", HH, MM, SS);
    }

    /// <summary>
    /// 將DataTable畫成表格
    /// </summary>
    /// <param name="dt">DataTable</param>
    /// <returns>string</returns>
    public string drawtable(DataTable dt)
    {
        string str = "<table border='1' width='100%'>";
        str += "<tr><td colspan='" + dt.Columns.Count.ToString() + "'>" + dt.Rows.Count.ToString() + "</td></tr>";
        if (dt.Rows.Count > 0)
        {
            str += "<tr>";
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                str += "<td>" + dt.Columns[j].ToString() + "</td>";
            }
            str += "</tr>";
        }

        for (int i = 0; i < dt.Rows.Count; i++)
        {
            str += "<tr>";
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                str += "<td>" + dt.Rows[i][j].ToString() + "</td>";
            }
            //Response.Write(rstr);
            str += "</tr>";
        }
        str += "</table>";
        return str;
    }

    /// <summary>
    /// 日期 轉成特定格式
    /// </summary>
    /// <param name="dt">日期</param>
    /// <param name="fmt">格式yyyy/MM/dd hh:mm:ss</param>
    /// <returns>string</returns>
    public string dt2format(string dt, string fmt)
    {
        return DateTime.Parse(dt).ToString(fmt);
    }

    /// <summary>
    /// DataRow 轉成 DataTable
    /// </summary>
    /// <param name="dt">DataTable</param>
    /// <param name="condition">select 條件</param>
    /// <returns>DataTable</returns>
    public DataTable GetNewDataTable(DataTable dt, string condition)
    {
        DataTable ndt = new DataTable();
        ndt = dt.Clone();
        DataRow[] dr = dt.Select(condition);
        for (int i = 0; i <= dr.Length - 1; i++)
        {
            ndt.ImportRow((DataRow)dr[i]);
        }
        return ndt;
    }


    /// <summary>
    /// 取得BPM專案連線字串
    /// </summary>
    /// <returns>ConnString</returns>
    public string SqlConnStr()
    {
        const string userRoot = "HKEY_LOCAL_MACHINE";
        //const string subkey = "Software\\NewType\\AutoWeb.Net";
        //const string keyName = userRoot + "\\" + subkey;
        string keyName = userRoot + "\\" + subkey;
        string Path = (string)Registry.GetValue(keyName, "Root", -1);
        Path = Path.Replace("\\", "\\\\");
        XdbcInfoIO objXdbc = new XdbcInfoIO();
        string FileName = Path + "\\\\Database\\\\Project\\\\" + project + "\\\\Connection\\\\" + bpmconn + ".xdbc.xmf";
        //string FileName = Path + "\\\\Database\\\\Project\\\\FlowMasterBPM\\\\BPM\\\\Connection\\\\FlowMasterBPM.xdbc.xmf";
        objXdbc.LoadFile(FileName, "");
        string connectionString = objXdbc.XdbcConnection.sOleDBConnectString;
        //OleDbConnection Conn = new OleDbConnection(connectionString);
        //Conn.Open();

        return connectionString;
    }

    /// <summary>
    /// 取得JB連線字串
    /// </summary>
    /// <returns>ConnString</returns>
    public string SqlConnStrJB()
    {
        const string userRoot = "HKEY_LOCAL_MACHINE";
        //const string subkey = "Software\\NewType\\AutoWeb.Net";
        //const string keyName = userRoot + "\\" + subkey;
        string keyName = userRoot + "\\" + subkey;
        string Path = (string)Registry.GetValue(keyName, "Root", -1);
        Path = Path.Replace("\\", "\\\\");
        XdbcInfoIO objXdbc = new XdbcInfoIO();
        string FileName = Path + "\\\\Database\\\\Project\\\\" + project + "\\\\Connection\\\\" + jbconn + ".xdbc.xmf";
        //string FileName = Path + "\\\\Database\\\\Project\\\\FlowMasterBPM\\\\BPM\\\\Connection\\\\jbHR.xdbc.xmf";
        objXdbc.LoadFile(FileName, "");
        string connectionString = objXdbc.XdbcConnection.sOleDBConnectString;
        //OleDbConnection Conn = new OleDbConnection(connectionString);
        //Conn.Open();

        return connectionString;
    }

    /// <summary>
    /// 執行SQL COMMAND
    /// </summary>
    /// <param name="sqlq">SQL COMMAND</param>
    /// <returns>DataTable</returns>
    public string ExecSqlQuery(string sqlq)
    {
        string str = "";
        using (OleDbConnection Conn = new OleDbConnection(SqlConnStr()))
        {
            OleDbCommand cmd = new OleDbCommand(sqlq, Conn);
            Conn.Open();
            OleDbDataReader dr = cmd.ExecuteReader();
            if (dr.Read()) str = dr[0].ToString();
            cmd.Clone();
            dr.Close();
            //Conn.Close();
        }
        return str;
    }

    /// <summary>
    /// 執行SQL COMMAND
    /// </summary>
    /// <param name="sqlq">SQL COMMAND</param>
    /// <returns>DataTable</returns>
    public void drawtable(string sqlq)
    {
        using (OleDbConnection Conn = new OleDbConnection(SqlConnStr()))
        {
            OleDbCommand cmd = new OleDbCommand(sqlq, Conn);
            Conn.Open();
            OleDbDataReader dr = cmd.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(dr);
            HttpContext.Current.Response.Write(drawtable(dt));
            cmd.Clone();
            dr.Close();
            //Conn.Close();
        }
        return;
    }

    /// <summary>
    /// 執行SQL COMMAND
    /// </summary>
    /// <param name="sqlq">SQL COMMAND</param>
    /// <returns>DataTable</returns>
    public string ExecSqlCommand(string sqlq)
    {
        string str = "";
        using (OleDbConnection Conn = new OleDbConnection(SqlConnStr()))
        {
            OleDbCommand cmd = new OleDbCommand(sqlq, Conn);
            Conn.Open();
            str = cmd.ExecuteNonQuery().ToString();
            cmd.Dispose();
            //Conn.Close();
        }
        return str;
    }

    /// <summary>
    /// 執行SQL COMMAND
    /// </summary>
    /// <param name="sqlq">SQL COMMAND</param>
    /// <returns>DataTable</returns>
    //public string ExecSqlCommandJB(string sqlq)
    //{
    //    string str = "";
    //    using (OleDbConnection Conn = new OleDbConnection(SqlConnStrJB()))
    //    {
    //        OleDbCommand cmd = new OleDbCommand(sqlq, Conn);
    //        Conn.Open();
    //        str = cmd.ExecuteNonQuery().ToString();
    //        cmd.Dispose();
    //        //Conn.Close();
    //    }
    //    return str;
    //}

    /// <summary>
    /// 執行SQL COMMAND Log
    /// </summary>
    /// <param name="isdo">1 do 0 no</param>
    /// <param name="purl">使用的網頁</param>
    /// <param name="prid">RequisitionID</param>
    /// <param name="jbcmd">呼叫的指令</param>
    /// <param name="jbrstr">呼叫的原始結果</param>
    /// <param name="wstr">呼叫的結果整理</param>
    /// <param name="cateid">jb func cateid</param>
    /// <returns>int result</returns>
    public string ExecSqlLog(int isdo, string purl, string prid, string jbcmd, string jbrstr, string wstr, string cateid)
    {
        string str = "";
        string sqlq1 = "";
        sqlq1 += " if OBJECT_ID('WHRJBdoLog') is null begin  ";
        sqlq1 += " create table WHRJBdoLog(UniqueID bigint identity(1,1),pdate datetime default(getdate()),purl nvarchar(max) ";
        sqlq1 += " ,prid varchar(100),jbcmd nvarchar(max),jbrstr nvarchar(max),wstr nvarchar(max),jbcate int,jbmemo nvarchar(max)); end  ";
        string sqlq = "insert into WHRJBdoLog(purl,prid,jbcmd,jbrstr,wstr,jbcate)values(?,?,?,?,?,?) ";
        if (isdo == 1)
        {
            using (OleDbConnection Conn = new OleDbConnection(SqlConnStr()))
            {
                /*
                 ASP.NET下OleDbCommand使用參數連接SQL Server
                 http://dqno1.org/dqno1discuz/thread-3110-1-1.html
                 
                 說明，由於OleDbCommand無法使用『@欄位名稱』 這樣的方式來傳遞參數，僅能使用問號『?』的型態來傳遞
                 因此在使用時，必須依照sql指令中問號的順序，依序使用Parameters.Add方法來新增
                 
                 conn.open()
                 dim dept,plan_name as string 
                 
                 dept = Request.Form("deptcode")
                 plan_name  = Request.Form("plan_name")
                 
                 
                 '注意：問號為參數使用方式，不是打錯字，也不是資料錯誤
                 sqlstr = "insert into apply(dept,plan_name) values(?,?)" 
                 
                 Dim cmd As New OleDbCommand(sqlstr, conn)
                 cmd.Parameters.Clear()
                 
                 ' 請務必依照順序dept-->plan_name填入
                 
                 '這裡的參數名稱可以使用問號，也可以直接如下使用：
                 'cmd.Parameters.Add(New OleDbParameter("單位名稱", dept))
                 'cmd.Parameters.Add(New OleDbParameter("計畫名稱", dept))
                 
                 cmd.Parameters.Add(New OleDbParameter("?", dept))
                 cmd.Parameters.Add(New OleDbParameter("?", plan_name))
                 
                 cmd.ExecuteNonQuery()
                 */
                OleDbCommand cmd = new OleDbCommand(sqlq1 + sqlq, Conn);
                cmd.Parameters.Clear();
                cmd.Parameters.Add(new OleDbParameter("?", purl));
                cmd.Parameters.Add(new OleDbParameter("?", prid));
                cmd.Parameters.Add(new OleDbParameter("?", jbcmd));
                cmd.Parameters.Add(new OleDbParameter("?", jbrstr));
                cmd.Parameters.Add(new OleDbParameter("?", wstr));
                cmd.Parameters.Add(new OleDbParameter("?", cateid));

                /*
                cmd.Parameters.Add("@purl",OleDbType.VarWChar);
                cmd.Parameters.Add("@prid", OleDbType.VarWChar);
                cmd.Parameters.Add("@jbcmd", OleDbType.VarWChar);
                cmd.Parameters.Add("@jbrstr", OleDbType.VarWChar);
                cmd.Parameters.Add("@wstr", OleDbType.VarWChar);
                cmd.Parameters["@purl"].Value = purl;
                cmd.Parameters["@prid"].Value = prid;
                cmd.Parameters["@jbcmd"].Value = jbcmd;
                cmd.Parameters["@jbrstr"].Value = jbrstr;
                cmd.Parameters["@wstr"].Value = wstr;
                */
                Conn.Open();
                str = cmd.ExecuteNonQuery().ToString();
                cmd.Dispose();
                //Conn.Close();
            }
        }
        return str;
    }

    /// <summary>
    /// 執行ExecSqlQuery
    /// </summary>
    /// <param name="sqlq">SQL COMMAND</param>
    /// <param name="rVar">欄位</param>
    /// <returns>DataTable</returns>
    public string ExecSqlQuery(string sqlq, string rVar)
    {
        string str = "";
        using (OleDbConnection Conn = new OleDbConnection(SqlConnStr()))
        {
            OleDbCommand cmd = new OleDbCommand(sqlq, Conn);
            Conn.Open();
            OleDbDataReader dr = cmd.ExecuteReader();
            if (dr.Read()) str = dr[0].ToString();
            cmd.Dispose();
            dr.Close();
            //Conn.Close();
        }
        return str;
    }

    /*
    public string ExecSqlQuery(string sqlq, string rVar)
    {
        string str = "";
        //using (SqlConnection Conn = new SqlConnection("Password=newtype;Persist Security Info=True;User ID=sa;Initial Catalog=FlowMasterBPM;Data Source=."))
        {
            //SqlConnection Conn = new SqlConnection("Provider=SQLOLEDB.1;Password=newtype;Persist Security Info=True;User ID=sa;Initial Catalog=FlowMasterBPM;Data Source=.");
            //Conn.Open();
            //string sqlq = "select DATEDIFF(s,'1970/1/1 00:00:00','2003/8/10 16:00:00') as c ";
            SqlCommand cmd = new SqlCommand(sqlq, Conn);
            Conn.Open();
            SqlDataReader dr = cmd.ExecuteReader();
            if (dr.Read()) str = dr[0].ToString();
            cmd.Dispose();
            dr.Close();
            //Conn.Close();
        }
        return str;
    }
    */



    /// <summary>
    /// 可以反加兩個日期之間任何一個時間單位。
    /// </summary>
    /// <param name="DateTime1">DateTime</param>
    /// <param name="DateTime2">DateTime</param>
    /// <returns>string</returns>
    public string DateDiff(DateTime DateTime1, DateTime DateTime2)
    {
        string dateDiff = null;
        TimeSpan ts1 = new TimeSpan(DateTime1.Ticks);
        TimeSpan ts2 = new TimeSpan(DateTime2.Ticks);
        TimeSpan ts = ts1.Subtract(ts2).Duration();
        dateDiff = ts.Days.ToString() + "天" + ts.Hours.ToString() + "小時" + ts.Minutes.ToString() + "分鐘" + ts.Seconds.ToString() + "秒";
        return dateDiff;
    }
    /*
    IT資訊之家 http://www.it55.com
	
	說明：
	1.DateTime值類型代表了一個從公元0001年1月1日0點0分0秒到公元9999年12月31日23點59分59秒之間的具體体日期時刻。
	因此，你可以用DateTime值類型來描述任何在想像范圍之內的時間。一個DateTime值代表了一個具體的時刻
	2.TimeSpan值包含了許多屬性與方法，用于訪問或處理一個TimeSpan值
	下面的列表涵蓋了其中的一部分：
	Add：與另一個TimeSpan值相加。
	Days:返回用天數計算的TimeSpan值。
	Duration:獲取TimeSpan的絕對值。
	Hours:返回用小時計算的TimeSpan值
	Milliseconds:返回用毫秒計算的TimeSpan值。
	Minutes:返回用分鐘計算的TimeSpan值。
	Negate:返回當前實例的相反數。
	Seconds:返回用秒計算的TimeSpan值。
	Subtract:從中減去另一個TimeSpan值。
	Ticks:返回TimeSpan值的tick數。
	TotalDays:返回TimeSpan值表示的天數。
	TotalHours:返回TimeSpan值表示的小時數。
	TotalMilliseconds:返回TimeSpan值表示的毫秒數。
	TotalMinutes:返回TimeSpan值表示的分鐘數。
	TotalSeconds:返回TimeSpan值表示的秒數。
    */


    /// <summary>
    /// 是否為上班日 True = 要計算(包含假別是否要包含假日) sNobr>工號,dDate>日期,sHcode>假別 
    /// </summary>
    /// <param name="sNobr">sNobr</param>
    /// <param name="sDate">sDate</param>
    /// <param name="sHcode">sHcode</param>
    /// <returns>bool</returns>
    //public bool IsWorkDayByAbs(string sNobr, string sDate, string sHcode)
    //{
    //    return ws.IsWorkDayByAbs(sNobr, sDate, sHcode);
    //}

    /// <summary>
    /// [逾期申請]—當申請日期減請假日期(迄)大於三個工作天(參考HR行事曆)時自動勾選 sDate>請假日期(迄)、加班日期 nDate>申請日期
    /// </summary>
    /// <param name="sDate">請假日期(迄)、加班日期</param>
    /// <param name="nDate">申請日期</param>
    /// <param name="sNobr">sNobr</param>
    /// <param name="sHcode">sHcode</param>
    /// <param name="cateid">cateid</param>
    /// <param name="prid">prid</param>
    /// <returns>int</returns>
    //public int GetWorkDayCnt(string sDate, string nDate, string sNobr, string sHcode, int cateid, string prid)
    //{
    //    int c = 0;
    //    if (sDate == "" || nDate == "" || sNobr == "" || sHcode == "") return -99;

    //    DateTime dt1, dt2;
    //    if (!DateTime.TryParse(sDate, out dt1)) return -1;
    //    if (!DateTime.TryParse(nDate, out dt2)) return -2;

    //    //DateTime dt1 = DateTime.Parse(sDate);
    //    //DateTime dt2 = DateTime.Parse(nDate);
    //    DateTime dtmp = dt1;
    //    /*
    //    小於零     這個執行個體早於 value。
    //    Zero       這個執行個體和 value 相同。
    //    大於零     這個執行個體晚於 value，或者 value 是 null。
    //     */
    //    //dt2 < dt1
    //    if (dt2.CompareTo(dt1) < 0) return c;
    //    else
    //    {
    //        //while (dtmp <= dt2)
    //        while (dtmp.CompareTo(dt2) <= 0)
    //        {
    //            if (IsWorkDayByAbs(sNobr, dtmp.ToString("yyyy/MM/dd"), sHcode)) c++;

    //            strjbcmd = "IsWorkDayByAbs(" + sNobr + "," + dtmp.ToString("yyyy/MM/dd") + "," + sHcode + ")";
    //            ExecSqlLog(dolog, "GetWorkDayCnt()", prid, strjbcmd, IsWorkDayByAbs(sNobr, dtmp.ToString("yyyy/MM/dd"), sHcode).ToString(), c.ToString(), "41");

    //            dtmp = dtmp.AddDays(1);
    //        }
    //    }
    //    return c;
    //}


    /// <summary>
    /// 取得 Request 並給初始值
    /// </summary>
    /// <param name="var">變數</param>
    /// <param name="initv">初始值</param>
    /// <returns>string</returns>
    public string getRequest(string var, string initv)
    {
        string rV = "";
        var Request = HttpContext.Current.Request;
        if (Request[var] != null) rV = Request[var].ToString();
        if (rV == "") rV = initv;
        return rV;
    }

    /// <summary>
    /// writeline
    /// </summary>
    /// <param name="var">變數</param>
    /// <param name="initv">初始值</param>
    /// <returns>string</returns>
    public string writeline(string var)
    {
        string rV = "";
        var response = HttpContext.Current.Response;
        response.Write(var + "<br>");
        return rV;
    }

}