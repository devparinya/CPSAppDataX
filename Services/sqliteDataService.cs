using CPSAppData.Models;
using CpsDataApp.Models;
using CpsDataApp.Services;
using QueueAppManager.Service;
using System.Data;
using System.Data.SQLite;
using ZXing;

namespace CPSAppData.Services
{
    public class sqliteDataService
    {
        SQLiteConnection? connection;
        dateTimeHelper datehelper =  new dateTimeHelper();
        dataTableService dtService = new dataTableService();
        DataTable dtResult = new DataTable();
        securityService security_ = new securityService();
        string dbFileName = "secure_cpsdata.db";
        //string PathData = @"D:\PARINYA\DevOps\devparinya\CPSAppData\";//D:\DevOps\CPSAppData"; //"D:\PARINYA\DevOps\devparinya\CPSAppData\";
        string dbFilePath = string.Empty;

        #region InnitialData
        public sqliteDataService()
        {
          dbFilePath =  Path.Combine(ApplicationHelper.PathData, string.Format(@"Data\{0}", dbFileName));
        
            if (doCreateDBFileData(dbFilePath))
            {
                string connectionString = string.Format("Data Source={0};", dbFilePath);
                connection = new SQLiteConnection(connectionString);
            }
        }
        public bool doCreateDBFileData(string pathdata)
        {
            try
            {
                if (!File.Exists(pathdata))
                {
                    SQLiteConnection.CreateFile(pathdata);
                }
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            return true;
        }
        public void doinitTableCPSMaster()
        {
            // CPSMaterData definition
            string createTableQuery =
                @"CREATE TABLE IF NOT EXISTS CPSMaterData(
                Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT
                ,CaseID TEXT
                ,CardStatus TEXT
                ,LedNumber TEXT
                ,WorkNo TEXT
                ,Maxmonth INTEGER
                ,CardNo1 TEXT
                ,JudgmentAmnt1 REAL
                ,PrincipleAmnt1 REAL
                ,PayAfterJudgAmt1 REAL
                ,DeptAmnt1 REAL
                ,LastPayDate1 TEXT
                ,CardNo2 TEXT
                ,JudgmentAmnt2 REAL
                ,PrincipleAmnt2 REAL
                ,PayAfterJudgAmt2 REAL
                ,DeptAmnt2 REAL
                ,LastPayDate2 TEXT
                ,CardNo3 TEXT
                ,JudgmentAmnt3 REAL
                ,PrincipleAmnt3 REAL
                ,PayAfterJudgAmt3 REAL
                ,DeptAmnt3 REAL
                ,LastPayDate3 TEXT
                ,CardNo4 TEXT
                ,JudgmentAmnt4 REAL      
                ,PrincipleAmnt4 REAL
                ,PayAfterJudgAmt4 REAL
                ,DeptAmnt4 REAL
                ,LastPayDate4 TEXT
                ,CardNo5 TEXT
                ,JudgmentAmnt5 REAL
                ,PrincipleAmnt5 REAL
                ,PayAfterJudgAmt5 REAL
                ,DeptAmnt5 REAL
                ,LastPayDate5 TEXT
                ,CardNo6 TEXT
                ,JudgmentAmnt6 REAL
                ,PrincipleAmnt6 REAL
                ,PayAfterJudgAmt6 REAL
                ,DeptAmnt6 REAL
                ,LastPayDate6 TEXT
                ,CustomerName TEXT
                ,CustomerID TEXT
                ,CustomerTel TEXT
                ,LegalStatus TEXT
                ,BlackNo TEXT
                ,RedNo TEXT
                ,JudgeDate TEXT
                ,CourtName TEXT
                ,LegalExecRemark TEXT
                ,LegalExecDate TEXT
                ,CollectorName TEXT
                ,CollectorTeam TEXT
                ,CollectorTel TEXT
                ,CustomFlag TEXT);";
            if (connection != null)
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand(createTableQuery, connection);
                command.ExecuteNonQuery();
                connection.Close();
            }
        }
        public void doinitTableDataCPSMaster()
        {
            string createTableQuery =
                @"CREATE TABLE IF NOT EXISTS DataCPSMaster(
                Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT
                ,CaseID TEXT
                ,CardStatus TEXT
                ,ListNo INTEGER
                ,LedNumber TEXT
                ,WorkNo TEXT
                ,CardNo TEXT
                ,JudgmentAmnt REAL
                ,PrincipleAmnt REAL
                ,PayAfterJudgAmt REAL
                ,DeptAmnt REAL
                ,LastPayDate TEXT               
                ,CustomerName TEXT
                ,CustomerID TEXT
                ,CustomerTel TEXT
                ,LegalStatus TEXT
                ,BlackNo TEXT
                ,RedNo TEXT
                ,JudgeDate TEXT
                ,CourtName TEXT
                ,LegalExecRemark TEXT
                ,LegalExecDate TEXT
                ,CollectorName TEXT
                ,CollectorTeam TEXT
                ,CollectorTel TEXT
                ,CustomFlag TEXT);";
            if (connection != null)
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand(createTableQuery, connection);
                command.ExecuteNonQuery();
                connection.Close();
            }
        }
        public void doinitTableMapData( )
        {
            string str_cmd = @"CREATE TABLE IF NOT EXISTS CPSMapping (
                            Id INTEGER NOT NULL UNIQUE,
                            MapName  TEXT,
                            CtrlName  TEXT,
                            DataColName    TEXT,
                            PRIMARY KEY(Id AUTOINCREMENT));";

            doExeCuteNonQueryCMD(str_cmd);
        }
        public void doinitTableCPSFestData()
        {
            string str_cmd = @"CREATE TABLE IF NOT EXISTS CPSFestData (
                            Id INTEGER NOT NULL UNIQUE,
                            CustomerID  TEXT,
                            CustomerName  TEXT,
                            LedNumber  TEXT,
                            WorkNo    TEXT,
                            LegalExecRemark TEXT,
                            PRIMARY KEY(Id AUTOINCREMENT));";

            doExeCuteNonQueryCMD(str_cmd);
        }
        public void initTableCPSPayData()
        {
            string str_cmd = @"CREATE TABLE IF NOT EXISTS CPSPayAmnt (
                            Id INTEGER NOT NULL UNIQUE,
                            CustomerID  TEXT,
                            CaseID  TEXT,
                            CustomerName  TEXT,
                            WorkNo  TEXT,
                            LedNumber  TEXT,
                            CardNo1  TEXT,
                            AccCloseAmnt1 REAL,  
                            AccClose6Amnt1  REAL,   
                            Installment6Amnt1 REAL, 
                            AccClose12Amnt1  REAL, 
                            Installment12Amnt1 REAL, 
                            AccClose24Amnt1   REAL,
                            Installment24Amnt1 REAL, 
                            CardNo2  TEXT,
                            AccCloseAmnt2 REAL,  
                            AccClose6Amnt2  REAL,
                            Installment6Amnt2 REAL, 
                            AccClose12Amnt2  REAL, 
                            Installment12Amnt2 REAL,
                            AccClose24Amnt2   REAL,
                            Installment24Amnt2 REAL,
                            CardNo3  TEXT,
                            AccCloseAmnt3 REAL,  
                            AccClose6Amnt3  REAL, 
                            Installment6Amnt3 REAL,
                            AccClose12Amnt3  REAL, 
                            Installment12Amnt3 REAL,
                            AccClose24Amnt3   REAL,
                            Installment24Amnt3 REAL,
                            CardNo4  TEXT,
                            AccCloseAmnt4 REAL,  
                            AccClose6Amnt4  REAL,  
                            Installment6Amnt4 REAL,
                            AccClose12Amnt4  REAL, 
                            Installment12Amnt4 REAL,
                            AccClose24Amnt4   REAL,
                            Installment24Amnt4 REAL,
                            CardNo5  TEXT,
                            AccCloseAmnt5 REAL,  
                            AccClose6Amnt5  REAL,  
                            Installment6Amnt5 REAL,
                            AccClose12Amnt5  REAL, 
                            Installment12Amnt5 REAL,
                            AccClose24Amnt5   REAL,
                            Installment24Amnt5 REAL,
                            CardNo6  TEXT,
                            AccCloseAmnt6 REAL,  
                            AccClose6Amnt6  REAL,  
                            Installment6Amnt6 REAL,
                            AccClose12Amnt6  REAL, 
                            Installment12Amnt6 REAL,
                            AccClose24Amnt6   REAL,
                            Installment24Amnt6 REAL,
                            LegalExecRemark TEXT,
                            PRIMARY KEY(Id AUTOINCREMENT));";
            doExeCuteNonQueryCMD(str_cmd);
        }
        public void doinitTableSettingData()
        {
        string str_cmd = @"CREATE TABLE IF NOT EXISTS SettingData (
                            Id INTEGER NOT NULL UNIQUE,
                            C2PathPdf  TEXT,
                            TablePathPdf  TEXT,
                            FestNo  TEXT,
                            FestName  TEXT,
                            FestDate  TEXT,
                            DateAtCalulate  TEXT,
                            FistDateInstall TEXT,
                            MaxMonth INTEGER,
                            PRIMARY KEY(Id));";

            doExeCuteNonQueryCMD(str_cmd);
            int count = checkExistSetting();
            if (count == 0)
            {
                doExeCuteNonQueryCMD("INSERT INTO SettingData (Id) VALUES (1);");
            }
        }
        #endregion
        public SettingData doLoadSettingData()
        {
            SettingData setData = new SettingData();
            string str_cmd = "SELECT " +
                "ifnull(C2PathPdf,'') as C2PathPdf" +
                ",ifnull(TablePathPdf,'') as TablePathPdf" +
                ",ifnull(FestNo,'') as FestNo" +
                ",ifnull(FestName,'') as FestName" +
                ",ifnull(FestDate,'') as FestDate" +
                ",ifnull(DateAtCalulate,'') as DateAtCalulate" +
                ",ifnull(FistDateInstall,'') as FistDateInstall" +
                ",ifnull(MaxMonth,0) as MaxMonth" +
                " FROM SettingData;";            
            if (connection != null)
            {
                connection.Open();
                using (SQLiteCommand command = new SQLiteCommand(str_cmd, connection))
                {
                    using (SQLiteDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                                setData.C2PathPdf = reader.GetString(0);
                                setData.TablePathPdf = reader.GetString(1);
                                setData.FestNo = reader.GetString(2);
                                setData.FestName = reader.GetString(3);
                                setData.FestDate = reader.GetString(4);
                                setData.DateAtCalulate = reader.GetString(5);
                                setData.FistDateInstall = reader.GetString(6);
                                setData.MaxMonth = reader.GetInt32(7);
                        }
                    }
                }
                connection.Close();
            }
            return setData;
        }
        private int checkExistSetting()
        {
            return doExeCuteScalarQueryCMD("SELECT Count(Id) FROM SettingData");
        }
        private int doExeCuteNonQueryCMD(string cmdstr)
        {
            int i_result = 0;
            if (connection != null)
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand(cmdstr, connection);
                i_result = command.ExecuteNonQuery();
                connection.Close();              
            }
            return i_result;
        }        
        private int doExeCuteScalarQueryCMD(string cmdstr)
        {
            int i_result = -1;
            if (connection != null)
            { 
                connection.Open();
                SQLiteCommand command = new SQLiteCommand(cmdstr, connection);
                var result = command.ExecuteScalar();
                if(result != null)
                {
                    int.TryParse(result.ToString(), out i_result);
                }
                connection.Close();
            }
            return i_result;
        }
        public void doUpdateSetingData(SettingData datasetting)
        {
            if (datasetting != null)
            {
                string str_update = string.Format("UPDATE SettingData " +
                    "SET C2PathPdf = '{0}', " +
                    "TablePathPdf = '{1}', " +
                    "FestNo = '{2}'," +
                    "FestName = '{3}'," +
                    "FestDate = '{4}'," +
                    "DateAtCalulate = '{5}'," +
                    "FistDateInstall = '{6}'," +
                    "MaxMonth = {7}"+
                    " WHERE Id =1;",datasetting.C2PathPdf,datasetting.TablePathPdf,datasetting.FestNo,datasetting.FestName,datasetting.FestDate,datasetting.DateAtCalulate, datasetting.FistDateInstall, datasetting.MaxMonth);

               int reult =  doExeCuteNonQueryCMD(str_update);
            }                                           
        }                                                      
        #region MasterMap MAPPING
        public bool  doSaveMapMaterCPS(string mapname, ComboBox[] ctrlAll)
        {
            int result = -1;
            if (connection != null)
            {
                doDeleteMapMater(mapname);
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                using (var cmd = connection.CreateCommand())
                {
                    cmd.CommandText = "INSERT INTO CPSMapping " +
                        "(MapName,CtrlName,DataColName) " +
                        "VALUES($MapName,$CtrlName,$DataColName);";

                    var param_MapName = cmd.CreateParameter();
                    param_MapName.ParameterName = "$MapName";
                    cmd.Parameters.Add(param_MapName);

                    var param_CtrlName = cmd.CreateParameter();
                    param_CtrlName.ParameterName = "$CtrlName";
                    cmd.Parameters.Add(param_CtrlName);

                    var param_DataColName = cmd.CreateParameter();
                    param_DataColName.ParameterName = "$DataColName";
                    cmd.Parameters.Add(param_DataColName);

                    for (int i = 0; i < ctrlAll.Length; i++)
                    {
                        param_MapName.Value = mapname;
                        param_CtrlName.Value = ctrlAll[i].Name;
                        param_DataColName.Value = ctrlAll[i].Text;
                        result = cmd.ExecuteNonQuery();
                    }
                    transaction.Commit();
                }
                connection.Close();
            }
            return result >= 0;
        }
        private void doDeleteMapMater(string MapName)
        {
            string str_cmd = string.Format(@"DELETE FROM CPSMapping WHERE MapName = '{0}';",MapName);
            if (connection != null)
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand(str_cmd, connection);
                command.ExecuteNonQuery();              
                connection.Close();
            }
        }
        public void doLoadMappingDate(ComboBox[] ctrlAll,string mapname)
        {
            foreach (ComboBox ctrl in ctrlAll)
            {
                ctrl.Text = doGetDataMAsterMAP(ctrl.Name, mapname);
            }
        }
        private string? doGetDataMAsterMAP(string ctrlname,string mapname)
        {
            string? mapnamecol = string.Empty;
            if (!string.IsNullOrEmpty(ctrlname))
            {
                if (connection != null)
                {                  
                    var sqlcmd = string.Format("SELECT DataColName FROM  CPSMapping WHERE MapName = '{1}' AND CtrlName ='{0}'", ctrlname, mapname);
                    try
                    {
                        connection.Open();
                        SQLiteCommand command = new SQLiteCommand(sqlcmd, connection);
                        var colmap_name = command.ExecuteScalar();
                        if (colmap_name == null)
                        {
                            mapnamecol = string.Empty;
                        }
                        else
                        {
                            mapnamecol = Convert.ToString(colmap_name);
                           
                        }
                        connection.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            return mapnamecol;
        }
        #endregion
        #region MasterCPS
        public List<DataCPSPerson> doCheckCPSFestDataNotInMaster()
        {
            List<DataCPSPerson> customdataall = new List<DataCPSPerson>();
            if (connection != null)
            {
                connection.Open();
                var sqlcmd = @" SELECT
                                ifnull(WorkNo,'') as WorkNo
                                ,ifnull(LedNumber,'') as LedNumber
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(CustomerName,'') as CustomerName
                            FROM CPSFestData
                            WHERE CustomerID NOT IN ( SELECT CustomerID FROM DataCPSMaster) 
                            ORDER by LedNumber; ";

                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        DataCPSPerson customdata = new DataCPSPerson();
                        customdata.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        customdata.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                        customdata.WorkNo = reader.GetString("WorkNo") ?? "";
                        customdata.LedNumber = reader.GetString("LedNumber") ?? "";
                        customdataall.Add(customdata);
                    }
                }
                connection.Close();
            }
            return customdataall;
        }
        public void doInsertMasterData(DataTable dataraw, ComboBox[] cmbcolctrl,bool clearmaster, ProgressBar progressLodedata)
        {
            if (dataraw != null) 
            {
                progressLodedata.Maximum = dataraw.Rows.Count;
                progressLodedata.Minimum = 0;
                progressLodedata.Step = 1;
                progressLodedata.Visible = true;
                progressLodedata.Value = 0;
                dtResult = dtService.doCreateResultDataTable();
                if (connection != null) 
                {
                    if (clearmaster)
                    {
                        doDeleteCPSMaterData();
                        doDeleteCPSCustomData();
                        doDeleteCPSFestData();
                    }

                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = @"INSERT INTO CPSMaterData (CaseID,CardStatus,CardNo1, JudgmentAmnt1, PrincipleAmnt1, PayAfterJudgAmt1, DeptAmnt1, LastPayDate1, CardNo2, JudgmentAmnt2, PrincipleAmnt2, PayAfterJudgAmt2, DeptAmnt2, LastPayDate2, 
                                 CardNo3, JudgmentAmnt3, PrincipleAmnt3, PayAfterJudgAmt3, DeptAmnt3, LastPayDate3, CardNo4, JudgmentAmnt4, PrincipleAmnt4, PayAfterJudgAmt4, DeptAmnt4, LastPayDate4, 
                                 CardNo5, JudgmentAmnt5, PrincipleAmnt5, PayAfterJudgAmt5, DeptAmnt5, LastPayDate5, CardNo6, JudgmentAmnt6, PrincipleAmnt6, PayAfterJudgAmt6, DeptAmnt6, LastPayDate6, 
                                 CustomerName, CustomerID, CustomerTel, LegalStatus, BlackNo, RedNo, JudgeDate, CourtName, LegalExecRemark, LegalExecDate, CollectorName, CollectorTeam, CollectorTel)
                                 VALUES($CaseID,$CardStatus,$CardNo1, $JudgmentAmnt1, $PrincipleAmnt1, $PayAfterJudgAmt1, $DeptAmnt1, $LastPayDate1, $CardNo2, $JudgmentAmnt2, $PrincipleAmnt2, $PayAfterJudgAmt2, $DeptAmnt2, $LastPayDate2,
                                 $CardNo3, $JudgmentAmnt3, $PrincipleAmnt3, $PayAfterJudgAmt3, $DeptAmnt3, $LastPayDate3, $CardNo4, $JudgmentAmnt4, $PrincipleAmnt4, $PayAfterJudgAmt4, $DeptAmnt4, $LastPayDate4,
                                 $CardNo5, $JudgmentAmnt5, $PrincipleAmnt5, $PayAfterJudgAmt5, $DeptAmnt5, $LastPayDate5, $CardNo6, $JudgmentAmnt6, $PrincipleAmnt6, $PayAfterJudgAmt6, $DeptAmnt6, $LastPayDate6,
                                 $CustomerName, $CustomerID,$CustomerTel, $LegalStatus, $BlackNo, $RedNo, $JudgeDate, $CourtName, $LegalExecRemark, $LegalExecDate, $CollectorName, $CollectorTeam, $CollectorTel);";

                        #region Parameter SET
                        var param_CaseID = cmd.CreateParameter();
                        param_CaseID.ParameterName = "$CaseID";
                        cmd.Parameters.Add(param_CaseID);

                        var param_CardStatus = cmd.CreateParameter();
                        param_CardStatus.ParameterName = "$CardStatus";
                        cmd.Parameters.Add(param_CardStatus);

                        var param_CardNo1 = cmd.CreateParameter();
                        param_CardNo1.ParameterName = "$CardNo1";
                        cmd.Parameters.Add(param_CardNo1);
                        var param_JudgmentAmnt1 = cmd.CreateParameter();
                        param_JudgmentAmnt1.ParameterName = "$JudgmentAmnt1";
                        cmd.Parameters.Add(param_JudgmentAmnt1);
                        var param_PrincipleAmnt1 = cmd.CreateParameter();
                        param_PrincipleAmnt1.ParameterName = "$PrincipleAmnt1";
                        cmd.Parameters.Add(param_PrincipleAmnt1);
                        var param_PayAfterJudgAmt1 = cmd.CreateParameter();
                        param_PayAfterJudgAmt1.ParameterName = "$PayAfterJudgAmt1";
                        cmd.Parameters.Add(param_PayAfterJudgAmt1);
                        var param_DeptAmnt1 = cmd.CreateParameter();
                        param_DeptAmnt1.ParameterName = "$DeptAmnt1";
                        cmd.Parameters.Add(param_DeptAmnt1);
                        var param_LastPayDate1 = cmd.CreateParameter();
                        param_LastPayDate1.ParameterName = "$LastPayDate1";
                        cmd.Parameters.Add(param_LastPayDate1);

                        var param_CardNo2 = cmd.CreateParameter();
                        param_CardNo2.ParameterName = "$CardNo2";
                        cmd.Parameters.Add(param_CardNo2);
                        var param_JudgmentAmnt2 = cmd.CreateParameter();
                        param_JudgmentAmnt2.ParameterName = "$JudgmentAmnt2";
                        cmd.Parameters.Add(param_JudgmentAmnt2);
                        var param_PrincipleAmnt2 = cmd.CreateParameter();
                        param_PrincipleAmnt2.ParameterName = "$PrincipleAmnt2";
                        cmd.Parameters.Add(param_PrincipleAmnt2);
                        var param_PayAfterJudgAmt2 = cmd.CreateParameter();
                        param_PayAfterJudgAmt2.ParameterName = "$PayAfterJudgAmt2";
                        cmd.Parameters.Add(param_PayAfterJudgAmt2);
                        var param_DeptAmnt2 = cmd.CreateParameter();
                        param_DeptAmnt2.ParameterName = "$DeptAmnt2";
                        cmd.Parameters.Add(param_DeptAmnt2);
                        var param_LastPayDate2 = cmd.CreateParameter();
                        param_LastPayDate2.ParameterName = "$LastPayDate2";
                        cmd.Parameters.Add(param_LastPayDate2);

                        var param_CardNo3 = cmd.CreateParameter();
                        param_CardNo3.ParameterName = "$CardNo3";
                        cmd.Parameters.Add(param_CardNo3);
                        var param_JudgmentAmnt3 = cmd.CreateParameter();
                        param_JudgmentAmnt3.ParameterName = "$JudgmentAmnt3";
                        cmd.Parameters.Add(param_JudgmentAmnt3);
                        var param_PrincipleAmnt3 = cmd.CreateParameter();
                        param_PrincipleAmnt3.ParameterName = "$PrincipleAmnt3";
                        cmd.Parameters.Add(param_PrincipleAmnt3);
                        var param_PayAfterJudgAmt3 = cmd.CreateParameter();
                        param_PayAfterJudgAmt3.ParameterName = "$PayAfterJudgAmt3";
                        cmd.Parameters.Add(param_PayAfterJudgAmt3);
                        var param_DeptAmnt3 = cmd.CreateParameter();
                        param_DeptAmnt3.ParameterName = "$DeptAmnt3";
                        cmd.Parameters.Add(param_DeptAmnt3);
                        var param_LastPayDate3 = cmd.CreateParameter();
                        param_LastPayDate3.ParameterName = "$LastPayDate3";
                        cmd.Parameters.Add(param_LastPayDate3);

                        var param_CardNo4 = cmd.CreateParameter();
                        param_CardNo4.ParameterName = "$CardNo4";
                        cmd.Parameters.Add(param_CardNo4);
                        var param_JudgmentAmnt4 = cmd.CreateParameter();
                        param_JudgmentAmnt4.ParameterName = "$JudgmentAmnt4";
                        cmd.Parameters.Add(param_JudgmentAmnt4);
                        var param_PrincipleAmnt4 = cmd.CreateParameter();
                        param_PrincipleAmnt4.ParameterName = "$PrincipleAmnt4";
                        cmd.Parameters.Add(param_PrincipleAmnt4);
                        var param_PayAfterJudgAmt4 = cmd.CreateParameter();
                        param_PayAfterJudgAmt4.ParameterName = "$PayAfterJudgAmt4";
                        cmd.Parameters.Add(param_PayAfterJudgAmt4);
                        var param_DeptAmnt4 = cmd.CreateParameter();
                        param_DeptAmnt4.ParameterName = "$DeptAmnt4";
                        cmd.Parameters.Add(param_DeptAmnt4);
                        var param_LastPayDate4 = cmd.CreateParameter();
                        param_LastPayDate4.ParameterName = "$LastPayDate4";
                        cmd.Parameters.Add(param_LastPayDate4);

                        var param_CardNo5 = cmd.CreateParameter();
                        param_CardNo5.ParameterName = "$CardNo5";
                        cmd.Parameters.Add(param_CardNo5);
                        var param_JudgmentAmnt5 = cmd.CreateParameter();
                        param_JudgmentAmnt5.ParameterName = "$JudgmentAmnt5";
                        cmd.Parameters.Add(param_JudgmentAmnt5);
                        var param_PrincipleAmnt5 = cmd.CreateParameter();
                        param_PrincipleAmnt5.ParameterName = "$PrincipleAmnt5";
                        cmd.Parameters.Add(param_PrincipleAmnt5);
                        var param_PayAfterJudgAmt5 = cmd.CreateParameter();
                        param_PayAfterJudgAmt5.ParameterName = "$PayAfterJudgAmt5";
                        cmd.Parameters.Add(param_PayAfterJudgAmt5);
                        var param_DeptAmnt5 = cmd.CreateParameter();
                        param_DeptAmnt5.ParameterName = "$DeptAmnt5";
                        cmd.Parameters.Add(param_DeptAmnt5);
                        var param_LastPayDate5 = cmd.CreateParameter();
                        param_LastPayDate5.ParameterName = "$LastPayDate5";
                        cmd.Parameters.Add(param_LastPayDate5);

                        var param_CardNo6 = cmd.CreateParameter();
                        param_CardNo6.ParameterName = "$CardNo6";
                        cmd.Parameters.Add(param_CardNo6);
                        var param_JudgmentAmnt6 = cmd.CreateParameter();
                        param_JudgmentAmnt6.ParameterName = "$JudgmentAmnt6";
                        cmd.Parameters.Add(param_JudgmentAmnt6);
                        var param_PrincipleAmnt6 = cmd.CreateParameter();
                        param_PrincipleAmnt6.ParameterName = "$PrincipleAmnt6";
                        cmd.Parameters.Add(param_PrincipleAmnt6);
                        var param_PayAfterJudgAmt6 = cmd.CreateParameter();
                        param_PayAfterJudgAmt6.ParameterName = "$PayAfterJudgAmt6";
                        cmd.Parameters.Add(param_PayAfterJudgAmt6);
                        var param_DeptAmnt6 = cmd.CreateParameter();
                        param_DeptAmnt6.ParameterName = "$DeptAmnt6";
                        cmd.Parameters.Add(param_DeptAmnt6);
                        var param_LastPayDate6 = cmd.CreateParameter();
                        param_LastPayDate6.ParameterName = "$LastPayDate6";
                        cmd.Parameters.Add(param_LastPayDate6);

                        var param_CustomerName = cmd.CreateParameter();
                        param_CustomerName.ParameterName = "$CustomerName";
                        cmd.Parameters.Add(param_CustomerName);

                        var param_CustomerID = cmd.CreateParameter();
                        param_CustomerID.ParameterName = "$CustomerID";
                        cmd.Parameters.Add(param_CustomerID);

                        var param_CustomerTel = cmd.CreateParameter();
                        param_CustomerTel.ParameterName = "$CustomerTel";
                        cmd.Parameters.Add(param_CustomerTel);

                        var param_LegalStatus = cmd.CreateParameter();
                        param_LegalStatus.ParameterName = "$LegalStatus";
                        cmd.Parameters.Add(param_LegalStatus);

                        var param_BlackNo = cmd.CreateParameter();
                        param_BlackNo.ParameterName = "$BlackNo";
                        cmd.Parameters.Add(param_BlackNo);

                        var param_RedNo = cmd.CreateParameter();
                        param_RedNo.ParameterName = "$RedNo";
                        cmd.Parameters.Add(param_RedNo);

                        var param_JudgeDate = cmd.CreateParameter();
                        param_JudgeDate.ParameterName = "$JudgeDate";
                        cmd.Parameters.Add(param_JudgeDate);

                        var param_CourtName = cmd.CreateParameter();
                        param_CourtName.ParameterName = "$CourtName";
                        cmd.Parameters.Add(param_CourtName);

                        var param_LegalExecRemark = cmd.CreateParameter();
                        param_LegalExecRemark.ParameterName = "$LegalExecRemark";
                        cmd.Parameters.Add(param_LegalExecRemark);

                        var param_LegalExecDate = cmd.CreateParameter();
                        param_LegalExecDate.ParameterName = "$LegalExecDate";
                        cmd.Parameters.Add(param_LegalExecDate);

                        var param_CollectorName = cmd.CreateParameter();
                        param_CollectorName.ParameterName = "$CollectorName";
                        cmd.Parameters.Add(param_CollectorName);

                        var param_CollectorTeam = cmd.CreateParameter();
                        param_CollectorTeam.ParameterName = "$CollectorTeam";
                        cmd.Parameters.Add(param_CollectorTeam);

                        var param_CollectorTel = cmd.CreateParameter();
                        param_CollectorTel.ParameterName = "$CollectorTel";
                        cmd.Parameters.Add(param_CollectorTel);
                        #endregion

                        for (int i = 0;i< dataraw.Rows.Count; i++)
                        {
                            param_CardNo1.Value = dataraw.Rows[i][cmbcolctrl[0].Text];
                            param_JudgmentAmnt1.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[1].Text].ToString())? 0: dataraw.Rows[i][cmbcolctrl[1].Text]);
                            param_PrincipleAmnt1.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[2].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[2].Text]);
                            param_PayAfterJudgAmt1.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[3].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[3].Text]);
                            param_DeptAmnt1.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[4].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[4].Text]);
                            param_LastPayDate1.Value = datehelper.ConverDateTODBStr(dataraw.Rows[i], cmbcolctrl[5].Text);// dataraw.Rows[i][cmbcolctrl[5].Text];
                            param_CardNo2.Value = dataraw.Rows[i][cmbcolctrl[6].Text]; 
                            param_JudgmentAmnt2.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[7].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[7].Text]);
                            param_PrincipleAmnt2.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[8].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[8].Text]);
                            param_PayAfterJudgAmt2.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[9].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[9].Text]);
                            param_DeptAmnt2.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[10].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[10].Text]);
                            param_LastPayDate2.Value = datehelper.ConverDateTODBStr(dataraw.Rows[i], cmbcolctrl[11].Text); //dataraw.Rows[i][cmbcolctrl[11].Text];
                            param_CardNo3.Value = dataraw.Rows[i][cmbcolctrl[12].Text];
                            param_JudgmentAmnt3.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[13].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[13].Text]);
                            param_PrincipleAmnt3.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[14].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[14].Text]);
                            param_PayAfterJudgAmt3.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[15].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[15].Text]);
                            param_DeptAmnt3.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[16].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[16].Text]);
                            param_LastPayDate3.Value = datehelper.ConverDateTODBStr(dataraw.Rows[i], cmbcolctrl[17].Text); //dataraw.Rows[i][cmbcolctrl[17].Text];
                            param_CardNo4.Value = dataraw.Rows[i][cmbcolctrl[18].Text];
                            param_JudgmentAmnt4.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[19].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[19].Text]);
                            param_PrincipleAmnt4.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[20].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[20].Text]);
                            param_PayAfterJudgAmt4.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[21].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[21].Text]);
                            param_DeptAmnt4.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[22].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[22].Text]);
                            param_LastPayDate4.Value = datehelper.ConverDateTODBStr(dataraw.Rows[i], cmbcolctrl[23].Text); //dataraw.Rows[i][cmbcolctrl[23].Text];
                            param_CardNo5.Value = dataraw.Rows[i][cmbcolctrl[24].Text];
                            param_JudgmentAmnt5.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[25].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[25].Text]);
                            param_PrincipleAmnt5.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[26].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[26].Text]);
                            param_PayAfterJudgAmt5.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[27].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[27].Text]);
                            param_DeptAmnt5.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[28].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[28].Text]);
                            param_LastPayDate5.Value = datehelper.ConverDateTODBStr(dataraw.Rows[i], cmbcolctrl[29].Text); //dataraw.Rows[i][cmbcolctrl[29].Text];
                            param_CardNo6.Value = dataraw.Rows[i][cmbcolctrl[30].Text];
                            param_JudgmentAmnt6.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[31].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[31].Text]);
                            param_PrincipleAmnt6.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[32].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[32].Text]);
                            param_PayAfterJudgAmt6.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[33].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[33].Text]);
                            param_DeptAmnt6.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[34].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[34].Text]);
                            param_LastPayDate6.Value = datehelper.ConverDateTODBStr(dataraw.Rows[i], cmbcolctrl[35].Text); //dataraw.Rows[i][cmbcolctrl[35].Text];
                            
                            string? customerid = Convert.ToString(dataraw.Rows[i][cmbcolctrl[37].Text] is null ? string.Empty : dataraw.Rows[i][cmbcolctrl[37].Text]);
                            string? customername = Convert.ToString(dataraw.Rows[i][cmbcolctrl[36].Text] is null ? string.Empty : dataraw.Rows[i][cmbcolctrl[36].Text]);
                            string? customertel = Convert.ToString(dataraw.Rows[i][cmbcolctrl[38].Text] is null ? string.Empty : dataraw.Rows[i][cmbcolctrl[38].Text]);

                            param_CustomerName.Value = security_.EncryptString(customername);
                            param_CustomerID.Value = security_.EncryptString(customerid);
                            param_CustomerTel.Value = security_.EncryptString(customertel);

                            param_LegalStatus.Value = dataraw.Rows[i][cmbcolctrl[39].Text];
                            param_BlackNo.Value = dataraw.Rows[i][cmbcolctrl[40].Text];
                            param_RedNo.Value = dataraw.Rows[i][cmbcolctrl[41].Text];
                            param_JudgeDate.Value = datehelper.ConverDateTODBStr(dataraw.Rows[i], cmbcolctrl[42].Text); //dataraw.Rows[i][cmbcolctrl[42].Text];
                            param_CourtName.Value = dataraw.Rows[i][cmbcolctrl[43].Text];                           
                            param_CollectorName.Value = dataraw.Rows[i][cmbcolctrl[44].Text];
                            param_CollectorTeam.Value = dataraw.Rows[i][cmbcolctrl[45].Text];
                            param_CollectorTel.Value = dataraw.Rows[i][cmbcolctrl[46].Text];
                            param_LegalExecDate.Value = datehelper.ConverDateTODBStr(dataraw.Rows[i], cmbcolctrl[47].Text); //dataraw.Rows[i][cmbcolctrl[47].Text];
                            param_LegalExecRemark.Value = dataraw.Rows[i][cmbcolctrl[48].Text];

                            param_CaseID.Value = Convert.ToString(dataraw.Rows[i][cmbcolctrl[49].Text]);
                            param_CardStatus.Value = dataraw.Rows[i][cmbcolctrl[50].Text];

                            int result_i = cmd.ExecuteNonQuery();                          
                            
                            if (result_i > 0)
                            {                                
                                if(!string.IsNullOrEmpty(customerid))doSetResultData(ref dtResult, customerid, "Y", "");
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(customerid)) doSetResultData(ref dtResult, customerid, "N", "");
                            }
                            progressLodedata.Value = i + 1;
                        }
                        transaction.Commit();
                    }
                    connection.Close();
                }
                progressLodedata.Visible = false;
            }
        }
        public void doInsertFestData(DataTable dataraw, ComboBox[] cmbcolctrl, ProgressBar progressLodedata,bool cleardata)
        {
            if (dataraw != null)
            {
                progressLodedata.Maximum = dataraw.Rows.Count;
                progressLodedata.Minimum = 0;
                progressLodedata.Step = 1;
                progressLodedata.Value = 0;
                progressLodedata.Visible = true;
                if (connection != null)
                {
                    if (cleardata)
                    {
                        doClearMasterWithFestDataClear();
                        doDeleteCPSFestData();
                    }
                    dtResult = dtService.doCreateResultDataTable();
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = @"INSERT INTO CPSFestData(CustomerID,CustomerName,WorkNo,LedNumber,LegalExecRemark) VALUES($CustomerID,$CustomerName,$WorkNo,$LedNumber,$LegalExecRemark);";

                        var param_LedNumber = cmd.CreateParameter();
                        param_LedNumber.ParameterName = "$LedNumber";
                        cmd.Parameters.Add(param_LedNumber);

                        var param_WorkNo = cmd.CreateParameter();
                        param_WorkNo.ParameterName = "$WorkNo";
                        cmd.Parameters.Add(param_WorkNo);

                        var param_CustomerID = cmd.CreateParameter();
                        param_CustomerID.ParameterName = "$CustomerID";
                        cmd.Parameters.Add(param_CustomerID);

                        var param_CustomerName = cmd.CreateParameter();
                        param_CustomerName.ParameterName = "$CustomerName";
                        cmd.Parameters.Add(param_CustomerName);

                        var param_LegalExecRemark = cmd.CreateParameter();
                        param_LegalExecRemark.ParameterName = "$LegalExecRemark";
                        cmd.Parameters.Add(param_LegalExecRemark);

                        for (int i = 0; i < dataraw.Rows.Count; i++)
                        {                           
                            string? customerid = Convert.ToString(dataraw.Rows[i][cmbcolctrl[0].Text] is null ? string.Empty : dataraw.Rows[i][cmbcolctrl[0].Text]);
                            string? customername = Convert.ToString(dataraw.Rows[1][cmbcolctrl[0].Text] is null ? string.Empty : dataraw.Rows[i][cmbcolctrl[1].Text]);

                            param_LedNumber.Value = Convert.ToString(dataraw.Rows[i][cmbcolctrl[3].Text]);
                            param_WorkNo.Value = Convert.ToString(dataraw.Rows[i][cmbcolctrl[2].Text]);
                            param_CustomerID.Value = security_.EncryptString(customerid);
                            param_CustomerName.Value = security_.EncryptString(customername);
                            param_LegalExecRemark.Value = dataraw.Rows[i][cmbcolctrl[4].Text];

                            int result_i = cmd.ExecuteNonQuery();
                           

                            if (result_i > 0) 
                            {

                                if(!string.IsNullOrEmpty(customerid)) doSetResultData(ref dtResult, customerid, "Y", "");                               
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(customerid)) doSetResultData(ref dtResult, customerid, "N", "");
                            }
                            progressLodedata.Value ++;
                        }
                        transaction.Commit();                                         
                    }
                    doUpdateFestData();
                    connection.Close();
                    progressLodedata.Visible = false;
                }

            }
        }
        public void doUpdateFestData() //New
        {
            string cmd_update = @"UPDATE DataCPSMaster
                                SET WorkNo = CPSFestData.WorkNo,LedNumber = CPSFestData.LedNumber,LegalExecRemark = CPSFestData.LegalExecRemark
                                FROM CPSFestData
                                WHERE (DataCPSMaster.CustomerID = CPSFestData.CustomerID)";
            if (connection != null)
            {
                if (connection.State == ConnectionState.Open)
                {
                    SQLiteCommand command = new SQLiteCommand(cmd_update, connection);
                    int reult = command.ExecuteNonQuery();
                }
            }
        }
        public void doUpdateCustomFlag()
        {
            string cmd_update = @"UPDATE DataCPSMaster
                                SET WorkNo = CPSPayAmnt.WorkNo,LedNumber = CPSPayAmnt.LedNumber,CustomFlag = 'Y',LegalExecRemark = CPSPayAmnt.LegalExecRemark
                                FROM CPSPayAmnt
                                WHERE (DataCPSMaster.CustomerID = CPSPayAmnt.CustomerID AND DataCPSMaster.CaseID = CPSPayAmnt.CaseID)";
            if (connection != null)
            {
                if (connection.State == ConnectionState.Open)
                {
                    SQLiteCommand command = new SQLiteCommand(cmd_update, connection);
                    int reult = command.ExecuteNonQuery();
                }
            }
        }
        private void doClearMasterWitCustomClear()//New
        {
            string cmd_update = @"UPDATE DataCPSMaster
                                SET WorkNo = '',LedNumber = '',CustomFlag = 'N',LegalExecRemark = ''
                                FROM CPSPayAmnt
                                WHERE (DataCPSMaster.CustomerID = CPSPayAmnt.CustomerID)";
            if (connection != null)
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand(cmd_update, connection);
                int reult = command.ExecuteNonQuery();
                connection.Close();
            }
        }
        public List<DataCPSPerson> doGetMasterDataCPSAll()
        {
            List<DataCPSPerson> cardall = new List<DataCPSPerson>();
            if (connection != null)
            {
                connection.Open();
                var sqlcmd = @"SELECT 
                                ifnull(CardNo1,'') as CardNo1
                                ,ifnull(JudgmentAmnt1,0) as JudgmentAmnt1
                                ,ifnull(PrincipleAmnt1,0) as PrincipleAmnt1
                                ,ifnull(PayAfterJudgAmt1,0) as PayAfterJudgAmt1
                                ,ifnull(DeptAmnt1,0) as DeptAmnt1
                                ,ifnull(LastPayDate1,'') as LastPayDate1
                                ,ifnull(CardNo2,'') as CardNo2
                                ,ifnull(JudgmentAmnt2,0) as JudgmentAmnt2
                                ,ifnull(PrincipleAmnt2,0) as PrincipleAmnt2
                                ,ifnull(PayAfterJudgAmt2,0) as PayAfterJudgAmt2
                                ,ifnull(DeptAmnt2,0) as DeptAmnt2
                                ,ifnull(LastPayDate2,'') as LastPayDate2
                                ,ifnull(CardNo3,'') as CardNo3
                                ,ifnull(JudgmentAmnt3,0) as JudgmentAmnt3
                                ,ifnull(PrincipleAmnt3,0) as PrincipleAmnt3
                                ,ifnull(PayAfterJudgAmt3,0) as PayAfterJudgAmt3
                                ,ifnull(DeptAmnt3,0) as DeptAmnt3
                                ,ifnull(LastPayDate3,'') as LastPayDate3                                
                                ,ifnull(CardNo4,'') as CardNo4
                                ,ifnull(JudgmentAmnt4,0) as JudgmentAmnt4
                                ,ifnull(PrincipleAmnt4,0) as PrincipleAmnt4
                                ,ifnull(PayAfterJudgAmt4,0) as PayAfterJudgAmt4
                                ,ifnull(DeptAmnt4,0) as DeptAmnt4
                                ,ifnull(LastPayDate4,'') as LastPayDate4
                                ,ifnull(CardNo5,'') as CardNo5
                                ,ifnull(JudgmentAmnt5,0) as JudgmentAmnt5
                                ,ifnull(PrincipleAmnt5,0) as PrincipleAmnt5
                                ,ifnull(PayAfterJudgAmt5,0) as PayAfterJudgAmt5
                                ,ifnull(DeptAmnt5,0) as DeptAmnt5
                                ,ifnull(LastPayDate5,'') as LastPayDate5
                                ,ifnull(CardNo6,'') as CardNo6
                                ,ifnull(JudgmentAmnt6,0) as JudgmentAmnt6
                                ,ifnull(PrincipleAmnt6,0) as PrincipleAmnt6
                                ,ifnull(PayAfterJudgAmt6,0) as PayAfterJudgAmt6
                                ,ifnull(DeptAmnt6,0) as DeptAmnt6
                                ,ifnull(LastPayDate6,'') as LastPayDate6
                                ,ifnull(CustomerName,'') as CustomerName
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(CustomerTel,'') as CustomerTel
                                ,ifnull(LegalStatus,'') as LegalStatus
                                ,ifnull(BlackNo,'') as BlackNo
                                ,ifnull(RedNo,'') as RedNo
                                ,ifnull(JudgeDate,'') as JudgeDate
                                ,ifnull(CourtName,'') as CourtName
                                ,ifnull(CollectorName,'') as CollectorName
                                ,ifnull(CollectorTeam,'') as CollectorTeam
                                ,ifnull(CollectorTel,'') as CollectorTel
                                ,ifnull(LegalExecDate,'') as LegalExecDate
                                ,ifnull(LegalExecRemark,'') as LegalExecRemark
                                ,ifnull(CustomFlag,'N') as CustomFlag
                               FROM CPSMaterData;";
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();
                connection.Close();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        DataCPSPerson carddat = new DataCPSPerson();
                        carddat.CardNo1 = reader.GetString("CardNo1") ?? "";
                        carddat.JudgmentAmnt1 = reader.GetDouble("JudgmentAmnt1");
                        carddat.PrincipleAmnt1 = reader.GetDouble("PrincipleAmnt1");
                        carddat.PayAfterJudgAmt1 = reader.GetDouble("PayAfterJudgAmt1");
                        carddat.DeptAmnt1 = reader.GetDouble("DeptAmnt1");
                        carddat.LastPayDate1 = reader.GetString("LastPayDate1") ?? "";

                        carddat.CardNo2 = reader.GetString("CardNo2") ?? "";
                        carddat.JudgmentAmnt2 = reader.GetDouble("JudgmentAmnt2");
                        carddat.PrincipleAmnt2 = reader.GetDouble("PrincipleAmnt2");
                        carddat.PayAfterJudgAmt2 = reader.GetDouble("PayAfterJudgAmt2");
                        carddat.DeptAmnt2 = reader.GetDouble("DeptAmnt2");
                        carddat.LastPayDate2 = reader.GetString("LastPayDate2") ?? "";

                        carddat.CardNo3 = reader.GetString("CardNo3") ?? "";
                        carddat.JudgmentAmnt3 = reader.GetDouble("JudgmentAmnt3");
                        carddat.PrincipleAmnt3 = reader.GetDouble("PrincipleAmnt3");
                        carddat.PayAfterJudgAmt3 = reader.GetDouble("PayAfterJudgAmt3");
                        carddat.DeptAmnt3 = reader.GetDouble("DeptAmnt3");
                        carddat.LastPayDate3 = reader.GetString("LastPayDate3") ?? "";

                        carddat.CardNo4 = reader.GetString("CardNo4") ?? "";
                        carddat.JudgmentAmnt4 = reader.GetDouble("JudgmentAmnt4");
                        carddat.PrincipleAmnt4 = reader.GetDouble("PrincipleAmnt4");
                        carddat.PayAfterJudgAmt4 = reader.GetDouble("PayAfterJudgAmt4");
                        carddat.DeptAmnt4 = reader.GetDouble("DeptAmnt4");
                        carddat.LastPayDate4 = reader.GetString("LastPayDate4") ?? "";

                        carddat.CardNo5 = reader.GetString("CardNo5") ?? "";
                        carddat.JudgmentAmnt5 = reader.GetDouble("JudgmentAmnt5");
                        carddat.PrincipleAmnt5 = reader.GetDouble("PrincipleAmnt5");
                        carddat.PayAfterJudgAmt5 = reader.GetDouble("PayAfterJudgAmt5");
                        carddat.DeptAmnt5 = reader.GetDouble("DeptAmnt5");
                        carddat.LastPayDate5 = reader.GetString("LastPayDate5") ?? "";

                        carddat.CardNo6 = reader.GetString("CardNo6") ?? "";
                        carddat.JudgmentAmnt6 = reader.GetDouble("JudgmentAmnt6");
                        carddat.PrincipleAmnt6 = reader.GetDouble("PrincipleAmnt6");
                        carddat.PayAfterJudgAmt6 = reader.GetDouble("PayAfterJudgAmt6");
                        carddat.DeptAmnt6 = reader.GetDouble("DeptAmnt6");
                        carddat.LastPayDate6 = reader.GetString("LastPayDate6") ?? "";

                        carddat.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                        carddat.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        carddat.CustomerTel = security_.DecryptString(reader.GetString("CustomerTel") ?? "");

                        carddat.LegalStatus = reader.GetString("LegalStatus") ?? "";
                        carddat.BlackNo = reader.GetString("BlackNo") ?? "";
                        carddat.RedNo = reader.GetString("RedNo") ?? "";
                        carddat.JudgeDate = reader.GetString("JudgeDate") ?? "";
                        carddat.CourtName = reader.GetString("CourtName") ?? "";
                        carddat.CollectorName = reader.GetString("CollectorName") ?? "";
                        carddat.CollectorTeam = reader.GetString("CollectorTeam") ?? "";
                        carddat.CollectorTel = reader.GetString("CollectorTel") ?? "";
                        carddat.LegalExecDate = reader.GetString("LegalExecDate") ?? "";
                        carddat.LegalExecRemark = reader.GetString("LegalExecRemark") ?? "";
                        carddat.CustomFlag = reader.GetString("CustomFlag") ?? "N";
                        cardall.Add(carddat);
                    }
                }
                
            }
                return cardall;
        }
        public List<DataCPSPerson> doGetMasterDuplicateIDDataCPSAll()
        {
            List<DataCPSPerson> cardall = new List<DataCPSPerson>();
            if (connection != null)
            {
                connection.Open();
                var sqlcmd = @"SELECT 
                                 ifnull(CaseID,'') as CaseID
                                ,ifnull(CardStatus,0) as CardStatus
                                ,ifnull(CardNo1,'') as CardNo1
                                ,ifnull(JudgmentAmnt1,0) as JudgmentAmnt1
                                ,ifnull(PrincipleAmnt1,0) as PrincipleAmnt1
                                ,ifnull(PayAfterJudgAmt1,0) as PayAfterJudgAmt1
                                ,ifnull(DeptAmnt1,0) as DeptAmnt1
                                ,ifnull(LastPayDate1,'') as LastPayDate1
                                ,ifnull(CardNo2,'') as CardNo2
                                ,ifnull(JudgmentAmnt2,0) as JudgmentAmnt2
                                ,ifnull(PrincipleAmnt2,0) as PrincipleAmnt2
                                ,ifnull(PayAfterJudgAmt2,0) as PayAfterJudgAmt2
                                ,ifnull(DeptAmnt2,0) as DeptAmnt2
                                ,ifnull(LastPayDate2,'') as LastPayDate2
                                ,ifnull(CardNo3,'') as CardNo3
                                ,ifnull(JudgmentAmnt3,0) as JudgmentAmnt3
                                ,ifnull(PrincipleAmnt3,0) as PrincipleAmnt3
                                ,ifnull(PayAfterJudgAmt3,0) as PayAfterJudgAmt3
                                ,ifnull(DeptAmnt3,0) as DeptAmnt3
                                ,ifnull(LastPayDate3,'') as LastPayDate3                                
                                ,ifnull(CardNo4,'') as CardNo4
                                ,ifnull(JudgmentAmnt4,0) as JudgmentAmnt4
                                ,ifnull(PrincipleAmnt4,0) as PrincipleAmnt4
                                ,ifnull(PayAfterJudgAmt4,0) as PayAfterJudgAmt4
                                ,ifnull(DeptAmnt4,0) as DeptAmnt4
                                ,ifnull(LastPayDate4,'') as LastPayDate4
                                ,ifnull(CardNo5,'') as CardNo5
                                ,ifnull(JudgmentAmnt5,0) as JudgmentAmnt5
                                ,ifnull(PrincipleAmnt5,0) as PrincipleAmnt5
                                ,ifnull(PayAfterJudgAmt5,0) as PayAfterJudgAmt5
                                ,ifnull(DeptAmnt5,0) as DeptAmnt5
                                ,ifnull(LastPayDate5,'') as LastPayDate5
                                ,ifnull(CardNo6,'') as CardNo6
                                ,ifnull(JudgmentAmnt6,0) as JudgmentAmnt6
                                ,ifnull(PrincipleAmnt6,0) as PrincipleAmnt6
                                ,ifnull(PayAfterJudgAmt6,0) as PayAfterJudgAmt6
                                ,ifnull(DeptAmnt6,0) as DeptAmnt6
                                ,ifnull(LastPayDate6,'') as LastPayDate6
                                ,ifnull(CustomerName,'') as CustomerName
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(CustomerTel,'') as CustomerTel
                                ,ifnull(LegalStatus,'') as LegalStatus
                                ,ifnull(BlackNo,'') as BlackNo
                                ,ifnull(RedNo,'') as RedNo
                                ,ifnull(JudgeDate,'') as JudgeDate
                                ,ifnull(CourtName,'') as CourtName
                                ,ifnull(CollectorName,'') as CollectorName
                                ,ifnull(CollectorTeam,'') as CollectorTeam
                                ,ifnull(CollectorTel,'') as CollectorTel
                                ,ifnull(LegalExecDate,'') as LegalExecDate
                                ,ifnull(LegalExecRemark,'') as LegalExecRemark
                                ,ifnull(CustomFlag,'N') as CustomFlag
                               FROM CPSMaterData
                               WHERE CustomerID in (SELECT CustomerID
                                                    FROM CPSMaterData
					                                GROUP BY CustomerID,CaseID
					                                HAVING COUNT(CustomerID) > 1)
                               ORDER by CustomerID;";
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();
                
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        DataCPSPerson carddat = new DataCPSPerson();
                        carddat.CaseID = reader.GetString("CaseID") ?? "";
                        carddat.CardStatus = reader.GetString("CardStatus") ?? "";
                        carddat.CardNo1 = reader.GetString("CardNo1") ?? "";
                        carddat.JudgmentAmnt1 = reader.GetDouble("JudgmentAmnt1");
                        carddat.PrincipleAmnt1 = reader.GetDouble("PrincipleAmnt1");
                        carddat.PayAfterJudgAmt1 = reader.GetDouble("PayAfterJudgAmt1");
                        carddat.DeptAmnt1 = reader.GetDouble("DeptAmnt1");
                        carddat.LastPayDate1 = reader.GetString("LastPayDate1") ?? "";

                        carddat.CardNo2 = reader.GetString("CardNo2") ?? "";
                        carddat.JudgmentAmnt2 = reader.GetDouble("JudgmentAmnt2");
                        carddat.PrincipleAmnt2 = reader.GetDouble("PrincipleAmnt2");
                        carddat.PayAfterJudgAmt2 = reader.GetDouble("PayAfterJudgAmt2");
                        carddat.DeptAmnt2 = reader.GetDouble("DeptAmnt2");
                        carddat.LastPayDate2 = reader.GetString("LastPayDate2") ?? "";

                        carddat.CardNo3 = reader.GetString("CardNo3") ?? "";
                        carddat.JudgmentAmnt3 = reader.GetDouble("JudgmentAmnt3");
                        carddat.PrincipleAmnt3 = reader.GetDouble("PrincipleAmnt3");
                        carddat.PayAfterJudgAmt3 = reader.GetDouble("PayAfterJudgAmt3");
                        carddat.DeptAmnt3 = reader.GetDouble("DeptAmnt3");
                        carddat.LastPayDate3 = reader.GetString("LastPayDate3") ?? "";

                        carddat.CardNo4 = reader.GetString("CardNo4") ?? "";
                        carddat.JudgmentAmnt4 = reader.GetDouble("JudgmentAmnt4");
                        carddat.PrincipleAmnt4 = reader.GetDouble("PrincipleAmnt4");
                        carddat.PayAfterJudgAmt4 = reader.GetDouble("PayAfterJudgAmt4");
                        carddat.DeptAmnt4 = reader.GetDouble("DeptAmnt4");
                        carddat.LastPayDate4 = reader.GetString("LastPayDate4") ?? "";

                        carddat.CardNo5 = reader.GetString("CardNo5") ?? "";
                        carddat.JudgmentAmnt5 = reader.GetDouble("JudgmentAmnt5");
                        carddat.PrincipleAmnt5 = reader.GetDouble("PrincipleAmnt5");
                        carddat.PayAfterJudgAmt5 = reader.GetDouble("PayAfterJudgAmt5");
                        carddat.DeptAmnt5 = reader.GetDouble("DeptAmnt5");
                        carddat.LastPayDate5 = reader.GetString("LastPayDate5") ?? "";

                        carddat.CardNo6 = reader.GetString("CardNo6") ?? "";
                        carddat.JudgmentAmnt6 = reader.GetDouble("JudgmentAmnt6");
                        carddat.PrincipleAmnt6 = reader.GetDouble("PrincipleAmnt6");
                        carddat.PayAfterJudgAmt6 = reader.GetDouble("PayAfterJudgAmt6");
                        carddat.DeptAmnt6 = reader.GetDouble("DeptAmnt6");
                        carddat.LastPayDate6 = reader.GetString("LastPayDate6") ?? "";

                        carddat.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                        carddat.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        carddat.CustomerTel = security_.DecryptString(reader.GetString("CustomerTel") ?? "");

                        carddat.LegalStatus = reader.GetString("LegalStatus") ?? "";
                        carddat.BlackNo = reader.GetString("BlackNo") ?? "";
                        carddat.RedNo = reader.GetString("RedNo") ?? "";
                        carddat.JudgeDate = reader.GetString("JudgeDate") ?? "";
                        carddat.CourtName = reader.GetString("CourtName") ?? "";
                        carddat.CollectorName = reader.GetString("CollectorName") ?? "";
                        carddat.CollectorTeam = reader.GetString("CollectorTeam") ?? "";
                        carddat.CollectorTel = reader.GetString("CollectorTel") ?? "";
                        carddat.LegalExecDate = reader.GetString("LegalExecDate") ?? "";
                        carddat.LegalExecRemark = reader.GetString("LegalExecRemark") ?? "";
                        carddat.CustomFlag = reader.GetString("CustomFlag") ?? "N";
                        cardall.Add(carddat);
                    }
                }
                connection.Close();
            }
            return cardall;
        }
        public List<DataCPSPerson> doGetDataCPSWithCustomerID(string customerid)
        {
            List<DataCPSPerson> cardall = new List<DataCPSPerson>();
            if (string.IsNullOrEmpty(customerid)) return cardall;

            string customerid_where = security_.EncryptString(customerid);
          
            if (connection != null)
            {
                connection.Open();
                var sqlcmd = string.Format(@"SELECT 
                                 ifnull(CaseID,'') as CaseID
                                ,ifnull(CardStatus,'') as CardStatus
                                ,ifnull(WorkNo,'') as WorkNo
                                ,ifnull(LedNumber,'') as LedNumber
                                ,ifnull(CardNo1,'') as CardNo1
                                ,ifnull(JudgmentAmnt1,0) as JudgmentAmnt1
                                ,ifnull(PrincipleAmnt1,0) as PrincipleAmnt1
                                ,ifnull(PayAfterJudgAmt1,0) as PayAfterJudgAmt1
                                ,ifnull(DeptAmnt1,0) as DeptAmnt1
                                ,ifnull(LastPayDate1,'') as LastPayDate1
                                ,ifnull(CardNo2,'') as CardNo2
                                ,ifnull(JudgmentAmnt2,0) as JudgmentAmnt2
                                ,ifnull(PrincipleAmnt2,0) as PrincipleAmnt2
                                ,ifnull(PayAfterJudgAmt2,0) as PayAfterJudgAmt2
                                ,ifnull(DeptAmnt2,0) as DeptAmnt2
                                ,ifnull(LastPayDate2,'') as LastPayDate2
                                ,ifnull(CardNo3,'') as CardNo3
                                ,ifnull(JudgmentAmnt3,0) as JudgmentAmnt3
                                ,ifnull(PrincipleAmnt3,0) as PrincipleAmnt3
                                ,ifnull(PayAfterJudgAmt3,0) as PayAfterJudgAmt3
                                ,ifnull(DeptAmnt3,0) as DeptAmnt3
                                ,ifnull(LastPayDate3,'') as LastPayDate3                                
                                ,ifnull(CardNo4,'') as CardNo4
                                ,ifnull(JudgmentAmnt4,0) as JudgmentAmnt4
                                ,ifnull(PrincipleAmnt4,0) as PrincipleAmnt4
                                ,ifnull(PayAfterJudgAmt4,0) as PayAfterJudgAmt4
                                ,ifnull(DeptAmnt4,0) as DeptAmnt4
                                ,ifnull(LastPayDate4,'') as LastPayDate4
                                ,ifnull(CardNo5,'') as CardNo5
                                ,ifnull(JudgmentAmnt5,0) as JudgmentAmnt5
                                ,ifnull(PrincipleAmnt5,0) as PrincipleAmnt5
                                ,ifnull(PayAfterJudgAmt5,0) as PayAfterJudgAmt5
                                ,ifnull(DeptAmnt5,0) as DeptAmnt5
                                ,ifnull(LastPayDate5,'') as LastPayDate5
                                ,ifnull(CardNo6,'') as CardNo6
                                ,ifnull(JudgmentAmnt6,0) as JudgmentAmnt6
                                ,ifnull(PrincipleAmnt6,0) as PrincipleAmnt6
                                ,ifnull(PayAfterJudgAmt6,0) as PayAfterJudgAmt6
                                ,ifnull(DeptAmnt6,0) as DeptAmnt6
                                ,ifnull(LastPayDate6,'') as LastPayDate6
                                ,ifnull(CustomerName,'') as CustomerName
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(CustomerTel,'') as CustomerTel
                                ,ifnull(LegalStatus,'') as LegalStatus
                                ,ifnull(BlackNo,'') as BlackNo
                                ,ifnull(RedNo,'') as RedNo
                                ,ifnull(JudgeDate,'') as JudgeDate
                                ,ifnull(CourtName,'') as CourtName
                                ,ifnull(CollectorName,'') as CollectorName
                                ,ifnull(CollectorTeam,'') as CollectorTeam
                                ,ifnull(CollectorTel,'') as CollectorTel
                                ,ifnull(LegalExecDate,'') as LegalExecDate
                                ,ifnull(LegalExecRemark,'') as LegalExecRemark
                                ,ifnull(CustomFlag,'N') as CustomFlag
                               FROM CPSMaterData
                               WHERE (CustomerID = '{0}');", customerid_where);
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();
                
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        DataCPSPerson carddat = new DataCPSPerson();
                        carddat.CaseID = reader.GetString("CaseID") ?? "";
                        carddat.CardStatus = reader.GetString("CardStatus") ?? "";
                        carddat.WorkNo = reader.GetString("WorkNo") ?? "";
                        carddat.LedNumber = reader.GetString("LedNumber") ?? "";
                        carddat.CardNo1 = reader.GetString("CardNo1") ?? "";
                        carddat.JudgmentAmnt1 = reader.GetDouble("JudgmentAmnt1");
                        carddat.PrincipleAmnt1 = reader.GetDouble("PrincipleAmnt1");
                        carddat.PayAfterJudgAmt1 = reader.GetDouble("PayAfterJudgAmt1");
                        carddat.DeptAmnt1 = reader.GetDouble("DeptAmnt1");
                        carddat.LastPayDate1 = reader.GetString("LastPayDate1") ?? "";

                        carddat.CardNo2 = reader.GetString("CardNo2") ?? "";
                        carddat.JudgmentAmnt2 = reader.GetDouble("JudgmentAmnt2");
                        carddat.PrincipleAmnt2 = reader.GetDouble("PrincipleAmnt2");
                        carddat.PayAfterJudgAmt2 = reader.GetDouble("PayAfterJudgAmt2");
                        carddat.DeptAmnt2 = reader.GetDouble("DeptAmnt2");
                        carddat.LastPayDate2 = reader.GetString("LastPayDate2") ?? "";

                        carddat.CardNo3 = reader.GetString("CardNo3") ?? "";
                        carddat.JudgmentAmnt3 = reader.GetDouble("JudgmentAmnt3");
                        carddat.PrincipleAmnt3 = reader.GetDouble("PrincipleAmnt3");
                        carddat.PayAfterJudgAmt3 = reader.GetDouble("PayAfterJudgAmt3");
                        carddat.DeptAmnt3 = reader.GetDouble("DeptAmnt3");
                        carddat.LastPayDate3 = reader.GetString("LastPayDate3") ?? "";

                        carddat.CardNo4 = reader.GetString("CardNo4") ?? "";
                        carddat.JudgmentAmnt4 = reader.GetDouble("JudgmentAmnt4");
                        carddat.PrincipleAmnt4 = reader.GetDouble("PrincipleAmnt4");
                        carddat.PayAfterJudgAmt4 = reader.GetDouble("PayAfterJudgAmt4");
                        carddat.DeptAmnt4 = reader.GetDouble("DeptAmnt4");
                        carddat.LastPayDate4 = reader.GetString("LastPayDate4") ?? "";

                        carddat.CardNo5 = reader.GetString("CardNo5") ?? "";
                        carddat.JudgmentAmnt5 = reader.GetDouble("JudgmentAmnt5");
                        carddat.PrincipleAmnt5 = reader.GetDouble("PrincipleAmnt5");
                        carddat.PayAfterJudgAmt5 = reader.GetDouble("PayAfterJudgAmt5");
                        carddat.DeptAmnt5 = reader.GetDouble("DeptAmnt5");
                        carddat.LastPayDate5 = reader.GetString("LastPayDate5") ?? "";

                        carddat.CardNo6 = reader.GetString("CardNo6") ?? "";
                        carddat.JudgmentAmnt6 = reader.GetDouble("JudgmentAmnt6");
                        carddat.PrincipleAmnt6 = reader.GetDouble("PrincipleAmnt6");
                        carddat.PayAfterJudgAmt6 = reader.GetDouble("PayAfterJudgAmt6");
                        carddat.DeptAmnt6 = reader.GetDouble("DeptAmnt6");
                        carddat.LastPayDate6 = reader.GetString("LastPayDate6") ?? "";

                        carddat.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                        carddat.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        carddat.CustomerTel = security_.DecryptString(reader.GetString("CustomerTel") ?? "");

                        carddat.LegalStatus = reader.GetString("LegalStatus") ?? "";
                        carddat.BlackNo = reader.GetString("BlackNo") ?? "";
                        carddat.RedNo = reader.GetString("RedNo") ?? "";
                        carddat.JudgeDate = reader.GetString("JudgeDate") ?? "";
                        carddat.CourtName = reader.GetString("CourtName") ?? "";
                        carddat.CollectorName = reader.GetString("CollectorName") ?? "";
                        carddat.CollectorTeam = reader.GetString("CollectorTeam") ?? "";
                        carddat.CollectorTel = reader.GetString("CollectorTel") ?? "";
                        carddat.LegalExecDate = reader.GetString("LegalExecDate") ?? "";
                        carddat.LegalExecRemark = reader.GetString("LegalExecRemark") ?? "";
                        carddat.CustomFlag = reader.GetString("CustomFlag") ?? "N";
                        
                        cardall.Add(carddat);
                    }
                }
                connection.Close();
            }
            return cardall;
        }
        public List<DataCPSPerson> doGetDataCPSWithCustomerIDAndCaseID(string customerid,string caseid)
        {
            List<DataCPSPerson> cardall = new List<DataCPSPerson>();
            if (string.IsNullOrEmpty(customerid)) return cardall;

            string customerid_where = security_.EncryptString(customerid);

            if (connection != null)
            {
                connection.Open();
                var sqlcmd = string.Format(@"SELECT 
                                 ifnull(CaseID,'') as CaseID
                                ,ifnull(CardStatus,'') as CardStatus
                                ,ifnull(WorkNo,'') as WorkNo
                                ,ifnull(LedNumber,'') as LedNumber
                                ,ifnull(CardNo1,'') as CardNo1
                                ,ifnull(JudgmentAmnt1,0) as JudgmentAmnt1
                                ,ifnull(PrincipleAmnt1,0) as PrincipleAmnt1
                                ,ifnull(PayAfterJudgAmt1,0) as PayAfterJudgAmt1
                                ,ifnull(DeptAmnt1,0) as DeptAmnt1
                                ,ifnull(LastPayDate1,'') as LastPayDate1
                                ,ifnull(CardNo2,'') as CardNo2
                                ,ifnull(JudgmentAmnt2,0) as JudgmentAmnt2
                                ,ifnull(PrincipleAmnt2,0) as PrincipleAmnt2
                                ,ifnull(PayAfterJudgAmt2,0) as PayAfterJudgAmt2
                                ,ifnull(DeptAmnt2,0) as DeptAmnt2
                                ,ifnull(LastPayDate2,'') as LastPayDate2
                                ,ifnull(CardNo3,'') as CardNo3
                                ,ifnull(JudgmentAmnt3,0) as JudgmentAmnt3
                                ,ifnull(PrincipleAmnt3,0) as PrincipleAmnt3
                                ,ifnull(PayAfterJudgAmt3,0) as PayAfterJudgAmt3
                                ,ifnull(DeptAmnt3,0) as DeptAmnt3
                                ,ifnull(LastPayDate3,'') as LastPayDate3                                
                                ,ifnull(CardNo4,'') as CardNo4
                                ,ifnull(JudgmentAmnt4,0) as JudgmentAmnt4
                                ,ifnull(PrincipleAmnt4,0) as PrincipleAmnt4
                                ,ifnull(PayAfterJudgAmt4,0) as PayAfterJudgAmt4
                                ,ifnull(DeptAmnt4,0) as DeptAmnt4
                                ,ifnull(LastPayDate4,'') as LastPayDate4
                                ,ifnull(CardNo5,'') as CardNo5
                                ,ifnull(JudgmentAmnt5,0) as JudgmentAmnt5
                                ,ifnull(PrincipleAmnt5,0) as PrincipleAmnt5
                                ,ifnull(PayAfterJudgAmt5,0) as PayAfterJudgAmt5
                                ,ifnull(DeptAmnt5,0) as DeptAmnt5
                                ,ifnull(LastPayDate5,'') as LastPayDate5
                                ,ifnull(CardNo6,'') as CardNo6
                                ,ifnull(JudgmentAmnt6,0) as JudgmentAmnt6
                                ,ifnull(PrincipleAmnt6,0) as PrincipleAmnt6
                                ,ifnull(PayAfterJudgAmt6,0) as PayAfterJudgAmt6
                                ,ifnull(DeptAmnt6,0) as DeptAmnt6
                                ,ifnull(LastPayDate6,'') as LastPayDate6
                                ,ifnull(CustomerName,'') as CustomerName
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(CustomerTel,'') as CustomerTel
                                ,ifnull(LegalStatus,'') as LegalStatus
                                ,ifnull(BlackNo,'') as BlackNo
                                ,ifnull(RedNo,'') as RedNo
                                ,ifnull(JudgeDate,'') as JudgeDate
                                ,ifnull(CourtName,'') as CourtName
                                ,ifnull(CollectorName,'') as CollectorName
                                ,ifnull(CollectorTeam,'') as CollectorTeam
                                ,ifnull(CollectorTel,'') as CollectorTel
                                ,ifnull(LegalExecDate,'') as LegalExecDate
                                ,ifnull(LegalExecRemark,'') as LegalExecRemark
                                ,ifnull(CustomFlag,'N') as CustomFlag
                               FROM CPSMaterData
                               WHERE (CustomerID = '{0}') And (CaseID = '{1}');", customerid_where,caseid);
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        DataCPSPerson carddat = new DataCPSPerson();
                        carddat.CaseID = reader.GetString("CaseID") ?? "";
                        carddat.CardStatus = reader.GetString("CardStatus") ?? "";
                        carddat.WorkNo = reader.GetString("WorkNo") ?? "";
                        carddat.LedNumber = reader.GetString("LedNumber") ?? "";
                        carddat.CardNo1 = reader.GetString("CardNo1") ?? "";
                        carddat.JudgmentAmnt1 = reader.GetDouble("JudgmentAmnt1");
                        carddat.PrincipleAmnt1 = reader.GetDouble("PrincipleAmnt1");
                        carddat.PayAfterJudgAmt1 = reader.GetDouble("PayAfterJudgAmt1");
                        carddat.DeptAmnt1 = reader.GetDouble("DeptAmnt1");
                        carddat.LastPayDate1 = reader.GetString("LastPayDate1") ?? "";

                        carddat.CardNo2 = reader.GetString("CardNo2") ?? "";
                        carddat.JudgmentAmnt2 = reader.GetDouble("JudgmentAmnt2");
                        carddat.PrincipleAmnt2 = reader.GetDouble("PrincipleAmnt2");
                        carddat.PayAfterJudgAmt2 = reader.GetDouble("PayAfterJudgAmt2");
                        carddat.DeptAmnt2 = reader.GetDouble("DeptAmnt2");
                        carddat.LastPayDate2 = reader.GetString("LastPayDate2") ?? "";

                        carddat.CardNo3 = reader.GetString("CardNo3") ?? "";
                        carddat.JudgmentAmnt3 = reader.GetDouble("JudgmentAmnt3");
                        carddat.PrincipleAmnt3 = reader.GetDouble("PrincipleAmnt3");
                        carddat.PayAfterJudgAmt3 = reader.GetDouble("PayAfterJudgAmt3");
                        carddat.DeptAmnt3 = reader.GetDouble("DeptAmnt3");
                        carddat.LastPayDate3 = reader.GetString("LastPayDate3") ?? "";

                        carddat.CardNo4 = reader.GetString("CardNo4") ?? "";
                        carddat.JudgmentAmnt4 = reader.GetDouble("JudgmentAmnt4");
                        carddat.PrincipleAmnt4 = reader.GetDouble("PrincipleAmnt4");
                        carddat.PayAfterJudgAmt4 = reader.GetDouble("PayAfterJudgAmt4");
                        carddat.DeptAmnt4 = reader.GetDouble("DeptAmnt4");
                        carddat.LastPayDate4 = reader.GetString("LastPayDate4") ?? "";

                        carddat.CardNo5 = reader.GetString("CardNo5") ?? "";
                        carddat.JudgmentAmnt5 = reader.GetDouble("JudgmentAmnt5");
                        carddat.PrincipleAmnt5 = reader.GetDouble("PrincipleAmnt5");
                        carddat.PayAfterJudgAmt5 = reader.GetDouble("PayAfterJudgAmt5");
                        carddat.DeptAmnt5 = reader.GetDouble("DeptAmnt5");
                        carddat.LastPayDate5 = reader.GetString("LastPayDate5") ?? "";

                        carddat.CardNo6 = reader.GetString("CardNo6") ?? "";
                        carddat.JudgmentAmnt6 = reader.GetDouble("JudgmentAmnt6");
                        carddat.PrincipleAmnt6 = reader.GetDouble("PrincipleAmnt6");
                        carddat.PayAfterJudgAmt6 = reader.GetDouble("PayAfterJudgAmt6");
                        carddat.DeptAmnt6 = reader.GetDouble("DeptAmnt6");
                        carddat.LastPayDate6 = reader.GetString("LastPayDate6") ?? "";

                        carddat.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                        carddat.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        carddat.CustomerTel = security_.DecryptString(reader.GetString("CustomerTel") ?? "");

                        carddat.LegalStatus = reader.GetString("LegalStatus") ?? "";
                        carddat.BlackNo = reader.GetString("BlackNo") ?? "";
                        carddat.RedNo = reader.GetString("RedNo") ?? "";
                        carddat.JudgeDate = reader.GetString("JudgeDate") ?? "";
                        carddat.CourtName = reader.GetString("CourtName") ?? "";
                        carddat.CollectorName = reader.GetString("CollectorName") ?? "";
                        carddat.CollectorTeam = reader.GetString("CollectorTeam") ?? "";
                        carddat.CollectorTel = reader.GetString("CollectorTel") ?? "";
                        carddat.LegalExecDate = reader.GetString("LegalExecDate") ?? "";
                        carddat.LegalExecRemark = reader.GetString("LegalExecRemark") ?? "";
                        carddat.CustomFlag = reader.GetString("CustomFlag") ?? "N";

                        cardall.Add(carddat);
                    }
                }
                connection.Close();
            }
            return cardall;
        }
        public List<DataCPSPerson> doGetDataCPSCustomerList(string customerid)
        {
            List<DataCPSPerson> cardall = new List<DataCPSPerson>();
            if (string.IsNullOrEmpty(customerid)) return cardall;
            string customerid_where = security_.EncryptString(customerid);

            if (connection != null)
            {
                connection.Open();
                var sqlcmd = string.Format(@"SELECT 
                                 ifnull(CaseID,'') as CaseID
                                ,ifnull(CardStatus,'') as CardStatus
                                ,ifnull(CustomerName,'') as CustomerName
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(LegalStatus,'') as LegalStatus
                               FROM CPSMaterData
                               WHERE (CustomerID = '{0}');", customerid_where);
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        DataCPSPerson carddat = new DataCPSPerson();
                        carddat.CaseID = reader.GetString("CaseID") ?? "";
                        carddat.CardStatus = reader.GetString("CardStatus") ?? "";
                        carddat.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                        carddat.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        carddat.LegalStatus = reader.GetString("LegalStatus") ?? "";
                        cardall.Add(carddat);
                    }
                }
                connection.Close();
            }
            return cardall;
        }
        #region New Dimanetion
        public List<DataCPSMaster> doGetDataCPSMasterByCustomerID(string customerid) //NEW
        {
            List<DataCPSMaster> dataMasterall= new List<DataCPSMaster>();
            if (string.IsNullOrEmpty(customerid)) return dataMasterall;
            string customerid_where = security_.EncryptString(customerid);

            if (connection != null)
            {
                connection.Open();
                var sqlcmd = string.Format(@"SELECT 
                                 ifnull(CaseID,'') as CaseID
                                ,ifnull(ListNo,0) as ListNo
                                ,ifnull(CardStatus,'') as CardStatus
                                ,ifnull(CustomerName,'') as CustomerName
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(LegalStatus,'') as LegalStatus
                                ,ifnull(LegalExecDate,'') as LegalExecDate
                                ,ifnull(LegalExecRemark,'') as LegalExecRemark
                               FROM DataCPSMaster
                               WHERE (CustomerID = '{0}');", customerid_where);
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();
                if (reader.HasRows)
                {                    
                    while (reader.Read())
                    {
                        string custID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        string caseID = reader.GetString("CaseID") ?? "";
                        int indexrow = -1;
                     
                        indexrow = dataMasterall.FindIndex(item => item.CustomerID == custID && item.CaseID == caseID);
                        if (indexrow < 0)
                        {
                            DataCPSMaster masterdata = new DataCPSMaster();
                            masterdata.CaseID = caseID;
                            masterdata.CardStatus = reader.GetString("CardStatus") ?? "";
                            masterdata.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                            masterdata.CustomerID = custID;
                            masterdata.LegalStatus = reader.GetString("LegalStatus") ?? "";
                            masterdata.LegalExecRemark = reader.GetString("LegalExecRemark") ?? "";
                            masterdata.LegalExecDate = datehelper.doGetShortDateTHFromDBToPDF(reader.GetString("LegalExecDate") ?? "");
                            dataMasterall.Add(masterdata);
                        }
                          
                    }
                }
                connection.Close();
            }
            return dataMasterall;
        }
        public void doGetDataCPSMasterByCaseID(ref List<DataCPSMaster> datamasterlist, string customerid,string caseid) //NEW
        {
            if (string.IsNullOrEmpty(caseid)) return;
            string customerid_where = security_.EncryptString(customerid);
            if (connection != null)
            {
                connection.Open();
                var sqlcmd = string.Format(@"SELECT 
                                 ifnull(CaseID,'') as CaseID
                                ,ifnull(ListNo,0) as ListNo
                                ,ifnull(CardStatus,'') as CardStatus
                                ,ifnull(CustomerName,'') as CustomerName
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(LegalStatus,'') as LegalStatus
                                ,ifnull(LegalExecDate,'') as LegalExecDate
                                ,ifnull(LegalExecRemark,'') as LegalExecRemark
                               FROM DataCPSMaster
                               WHERE (CaseID = '{0}')
                               AND (CustoMerID <> '{1}') ;", caseid,customerid);
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        string custID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        string caseID = reader.GetString("CaseID") ?? "";
                        int indexrow = -1;
                        indexrow = datamasterlist.FindIndex(item => item.CustomerID == custID && item.CaseID == caseid);                       
                        if (indexrow < 0)
                        {
                            DataCPSMaster masterdata = new DataCPSMaster();
                            masterdata.CaseID = reader.GetString("CaseID") ?? "";
                            masterdata.CardStatus = reader.GetString("CardStatus") ?? "";
                            masterdata.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                            masterdata.CustomerID = custID;
                            masterdata.LegalStatus = reader.GetString("LegalStatus") ?? "";
                            masterdata.LegalExecRemark = reader.GetString("LegalExecRemark") ?? "";
                            masterdata.LegalExecDate = datehelper.doGetShortDateTHFromDBToPDF(reader.GetString("LegalExecDate") ?? "");
                            datamasterlist.Add(masterdata);                           
                        }
                    }
                }
                connection.Close();
            }
        }
        public List<DataCPSMaster> doGetDataCPSMasterAllByCustomerID(string customerid, string caseid) //NEW
        {
            List<DataCPSMaster> masterall = new List<DataCPSMaster>();
            if (string.IsNullOrEmpty(customerid)) return masterall;

            string customerid_where = security_.EncryptString(customerid);
            string caseid_data = caseid;
            if (connection != null)
            {
                connection.Open();
                var sqlcmd = string.Format(@"SELECT 
                                 ifnull(CaseID,'') as CaseID
                                ,ifnull(CardStatus,'') as CardStatus
                                ,ifnull(WorkNo,'') as WorkNo
                                ,ifnull(LedNumber,'') as LedNumber
                                ,ifnull(CardNo,'') as CardNo
                                ,ifnull(JudgmentAmnt,0) as JudgmentAmnt
                                ,ifnull(PrincipleAmnt,0) as PrincipleAmnt
                                ,ifnull(PayAfterJudgAmt,0) as PayAfterJudgAmt
                                ,ifnull(DeptAmnt,0) as DeptAmnt
                                ,ifnull(LastPayDate,'') as LastPayDate
                                ,ifnull(CustomerName,'') as CustomerName
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(CustomerTel,'') as CustomerTel
                                ,ifnull(LegalStatus,'') as LegalStatus
                                ,ifnull(BlackNo,'') as BlackNo
                                ,ifnull(RedNo,'') as RedNo
                                ,ifnull(JudgeDate,'') as JudgeDate
                                ,ifnull(CourtName,'') as CourtName
                                ,ifnull(CollectorName,'') as CollectorName
                                ,ifnull(CollectorTeam,'') as CollectorTeam
                                ,ifnull(CollectorTel,'') as CollectorTel
                                ,ifnull(LegalExecDate,'') as LegalExecDate
                                ,ifnull(LegalExecRemark,'') as LegalExecRemark
                                ,ifnull(CustomFlag,'N') as CustomFlag
                               FROM DataCPSMaster
                               WHERE (CustomerID = '{0}') AND (CaseID = '{1}');", customerid_where, caseid_data);
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        DataCPSMaster carddat = new DataCPSMaster();
                        carddat.CaseID = reader.GetString("CaseID") ?? "";
                        carddat.CardStatus = reader.GetString("CardStatus") ?? "";
                        carddat.WorkNo = reader.GetString("WorkNo") ?? "";
                        carddat.LedNumber = reader.GetString("LedNumber") ?? "";
                        carddat.CardNo = reader.GetString("CardNo") ?? "";
                        carddat.JudgmentAmnt = reader.GetDouble("JudgmentAmnt");
                        carddat.PrincipleAmnt = reader.GetDouble("PrincipleAmnt");
                        carddat.PayAfterJudgAmt = reader.GetDouble("PayAfterJudgAmt");
                        carddat.DeptAmnt = reader.GetDouble("DeptAmnt");
                        carddat.LastPayDate = reader.GetString("LastPayDate") ?? "";

                        carddat.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                        carddat.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        carddat.CustomerTel = security_.DecryptString(reader.GetString("CustomerTel") ?? "");

                        carddat.LegalStatus = reader.GetString("LegalStatus") ?? "";
                        carddat.BlackNo = reader.GetString("BlackNo") ?? "";
                        carddat.RedNo = reader.GetString("RedNo") ?? "";
                        carddat.JudgeDate = reader.GetString("JudgeDate") ?? "";
                        carddat.CourtName = reader.GetString("CourtName") ?? "";
                        carddat.CollectorName = reader.GetString("CollectorName") ?? "";
                        carddat.CollectorTeam = reader.GetString("CollectorTeam") ?? "";
                        carddat.CollectorTel = reader.GetString("CollectorTel") ?? "";
                        carddat.LegalExecDate = reader.GetString("LegalExecDate") ?? "";
                        carddat.LegalExecRemark = reader.GetString("LegalExecRemark") ?? "";
                        carddat.CustomFlag = reader.GetString("CustomFlag") ?? "N";

                        masterall.Add(carddat);
                    }
                }
                connection.Close();
            }
            return masterall;
        }
        public List<DataCPSMaster> doGetDataCPSMasterAllByCaseID(ref List<DataCPSMaster> datamasterlist, string customerid,string caseid) //NEW
        {
            List<DataCPSMaster> masterall = new List<DataCPSMaster>();
            if (string.IsNullOrEmpty(customerid)) return masterall;

            string customerid_where = security_.EncryptString(customerid);
            if (connection != null)
            {
                connection.Open();
                var sqlcmd = string.Format(@"SELECT 
                                 ifnull(CaseID,'') as CaseID
                                ,ifnull(CardStatus,'') as CardStatus
                                ,ifnull(WorkNo,'') as WorkNo
                                ,ifnull(LedNumber,'') as LedNumber
                                ,ifnull(CardNo,'') as CardNo
                                ,ifnull(JudgmentAmnt,0) as JudgmentAmnt
                                ,ifnull(PrincipleAmnt,0) as PrincipleAmnt
                                ,ifnull(PayAfterJudgAmt,0) as PayAfterJudgAmt
                                ,ifnull(DeptAmnt,0) as DeptAmnt
                                ,ifnull(LastPayDate,'') as LastPayDate
                                ,ifnull(CustomerName,'') as CustomerName
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(CustomerTel,'') as CustomerTel
                                ,ifnull(LegalStatus,'') as LegalStatus
                                ,ifnull(BlackNo,'') as BlackNo
                                ,ifnull(RedNo,'') as RedNo
                                ,ifnull(JudgeDate,'') as JudgeDate
                                ,ifnull(CourtName,'') as CourtName
                                ,ifnull(CollectorName,'') as CollectorName
                                ,ifnull(CollectorTeam,'') as CollectorTeam
                                ,ifnull(CollectorTel,'') as CollectorTel
                                ,ifnull(LegalExecDate,'') as LegalExecDate
                                ,ifnull(LegalExecRemark,'') as LegalExecRemark
                                ,ifnull(CustomFlag,'N') as CustomFlag
                               FROM DataCPSMaster
                               WHERE (CaseID = '{0}') And (CustomerID <> '{1}' );", caseid,customerid_where);
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        DataCPSMaster carddat = new DataCPSMaster();
                        carddat.CaseID = reader.GetString("CaseID") ?? "";
                        carddat.CardStatus = reader.GetString("CardStatus") ?? "";
                        carddat.WorkNo = reader.GetString("WorkNo") ?? "";
                        carddat.LedNumber = reader.GetString("LedNumber") ?? "";
                        carddat.CardNo = reader.GetString("CardNo") ?? "";
                        carddat.JudgmentAmnt = reader.GetDouble("JudgmentAmnt");
                        carddat.PrincipleAmnt = reader.GetDouble("PrincipleAmnt");
                        carddat.PayAfterJudgAmt = reader.GetDouble("PayAfterJudgAmt");
                        carddat.DeptAmnt = reader.GetDouble("DeptAmnt");
                        carddat.LastPayDate = reader.GetString("LastPayDate") ?? "";

                        carddat.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                        carddat.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        carddat.CustomerTel = security_.DecryptString(reader.GetString("CustomerTel") ?? "");

                        carddat.LegalStatus = reader.GetString("LegalStatus") ?? "";
                        carddat.BlackNo = reader.GetString("BlackNo") ?? "";
                        carddat.RedNo = reader.GetString("RedNo") ?? "";
                        carddat.JudgeDate = reader.GetString("JudgeDate") ?? "";
                        carddat.CourtName = reader.GetString("CourtName") ?? "";
                        carddat.CollectorName = reader.GetString("CollectorName") ?? "";
                        carddat.CollectorTeam = reader.GetString("CollectorTeam") ?? "";
                        carddat.CollectorTel = reader.GetString("CollectorTel") ?? "";
                        carddat.LegalExecDate = reader.GetString("LegalExecDate") ?? "";
                        carddat.LegalExecRemark = reader.GetString("LegalExecRemark") ?? "";
                        carddat.CustomFlag = reader.GetString("CustomFlag") ?? "N";

                        masterall.Add(carddat);
                    }
                }
                connection.Close();
            }
            return masterall;
        }
        #endregion
        public List<DataCPSPerson> doGetDataCPSWithCaseID(string caseid)
        {
            List<DataCPSPerson> cardall = new List<DataCPSPerson>();
            if (string.IsNullOrEmpty(caseid)) return cardall;
            if (connection != null)
            {
                connection.Open();
                var sqlcmd = string.Format(@"SELECT 
                                 ifnull(CaseID,'') as CaseID
                                ,ifnull(CardStatus,'') as CardStatus
                                ,ifnull(WorkNo,'') as WorkNo
                                ,ifnull(LedNumber,'') as LedNumber
                                ,ifnull(CardNo1,'') as CardNo1
                                ,ifnull(JudgmentAmnt1,0) as JudgmentAmnt1
                                ,ifnull(PrincipleAmnt1,0) as PrincipleAmnt1
                                ,ifnull(PayAfterJudgAmt1,0) as PayAfterJudgAmt1
                                ,ifnull(DeptAmnt1,0) as DeptAmnt1
                                ,ifnull(LastPayDate1,'') as LastPayDate1
                                ,ifnull(CardNo2,'') as CardNo2
                                ,ifnull(JudgmentAmnt2,0) as JudgmentAmnt2
                                ,ifnull(PrincipleAmnt2,0) as PrincipleAmnt2
                                ,ifnull(PayAfterJudgAmt2,0) as PayAfterJudgAmt2
                                ,ifnull(DeptAmnt2,0) as DeptAmnt2
                                ,ifnull(LastPayDate2,'') as LastPayDate2
                                ,ifnull(CardNo3,'') as CardNo3
                                ,ifnull(JudgmentAmnt3,0) as JudgmentAmnt3
                                ,ifnull(PrincipleAmnt3,0) as PrincipleAmnt3
                                ,ifnull(PayAfterJudgAmt3,0) as PayAfterJudgAmt3
                                ,ifnull(DeptAmnt3,0) as DeptAmnt3
                                ,ifnull(LastPayDate3,'') as LastPayDate3                                
                                ,ifnull(CardNo4,'') as CardNo4
                                ,ifnull(JudgmentAmnt4,0) as JudgmentAmnt4
                                ,ifnull(PrincipleAmnt4,0) as PrincipleAmnt4
                                ,ifnull(PayAfterJudgAmt4,0) as PayAfterJudgAmt4
                                ,ifnull(DeptAmnt4,0) as DeptAmnt4
                                ,ifnull(LastPayDate4,'') as LastPayDate4
                                ,ifnull(CardNo5,'') as CardNo5
                                ,ifnull(JudgmentAmnt5,0) as JudgmentAmnt5
                                ,ifnull(PrincipleAmnt5,0) as PrincipleAmnt5
                                ,ifnull(PayAfterJudgAmt5,0) as PayAfterJudgAmt5
                                ,ifnull(DeptAmnt5,0) as DeptAmnt5
                                ,ifnull(LastPayDate5,'') as LastPayDate5
                                ,ifnull(CardNo6,'') as CardNo6
                                ,ifnull(JudgmentAmnt6,0) as JudgmentAmnt6
                                ,ifnull(PrincipleAmnt6,0) as PrincipleAmnt6
                                ,ifnull(PayAfterJudgAmt6,0) as PayAfterJudgAmt6
                                ,ifnull(DeptAmnt6,0) as DeptAmnt6
                                ,ifnull(LastPayDate6,'') as LastPayDate6
                                ,ifnull(CustomerName,'') as CustomerName
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(CustomerTel,'') as CustomerTel
                                ,ifnull(LegalStatus,'') as LegalStatus
                                ,ifnull(BlackNo,'') as BlackNo
                                ,ifnull(RedNo,'') as RedNo
                                ,ifnull(JudgeDate,'') as JudgeDate
                                ,ifnull(CourtName,'') as CourtName
                                ,ifnull(CollectorName,'') as CollectorName
                                ,ifnull(CollectorTeam,'') as CollectorTeam
                                ,ifnull(CollectorTel,'') as CollectorTel
                                ,ifnull(LegalExecDate,'') as LegalExecDate
                                ,ifnull(LegalExecRemark,'') as LegalExecRemark
                                ,ifnull(CustomFlag,'N') as CustomFlag
                               FROM CPSMaterData
                               WHERE (CaseID = '{0}');", caseid);
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        DataCPSPerson carddat = new DataCPSPerson();
                        carddat.CaseID = reader.GetString("CaseID") ?? "";
                        carddat.CardStatus = reader.GetString("CardStatus") ?? "";
                        carddat.WorkNo = reader.GetString("WorkNo") ?? "";
                        carddat.LedNumber = reader.GetString("LedNumber") ?? "";
                        carddat.CardNo1 = reader.GetString("CardNo1") ?? "";
                        carddat.JudgmentAmnt1 = reader.GetDouble("JudgmentAmnt1");
                        carddat.PrincipleAmnt1 = reader.GetDouble("PrincipleAmnt1");
                        carddat.PayAfterJudgAmt1 = reader.GetDouble("PayAfterJudgAmt1");
                        carddat.DeptAmnt1 = reader.GetDouble("DeptAmnt1");
                        carddat.LastPayDate1 = reader.GetString("LastPayDate1") ?? "";

                        carddat.CardNo2 = reader.GetString("CardNo2") ?? "";
                        carddat.JudgmentAmnt2 = reader.GetDouble("JudgmentAmnt2");
                        carddat.PrincipleAmnt2 = reader.GetDouble("PrincipleAmnt2");
                        carddat.PayAfterJudgAmt2 = reader.GetDouble("PayAfterJudgAmt2");
                        carddat.DeptAmnt2 = reader.GetDouble("DeptAmnt2");
                        carddat.LastPayDate2 = reader.GetString("LastPayDate2") ?? "";

                        carddat.CardNo3 = reader.GetString("CardNo3") ?? "";
                        carddat.JudgmentAmnt3 = reader.GetDouble("JudgmentAmnt3");
                        carddat.PrincipleAmnt3 = reader.GetDouble("PrincipleAmnt3");
                        carddat.PayAfterJudgAmt3 = reader.GetDouble("PayAfterJudgAmt3");
                        carddat.DeptAmnt3 = reader.GetDouble("DeptAmnt3");
                        carddat.LastPayDate3 = reader.GetString("LastPayDate3") ?? "";

                        carddat.CardNo4 = reader.GetString("CardNo4") ?? "";
                        carddat.JudgmentAmnt4 = reader.GetDouble("JudgmentAmnt4");
                        carddat.PrincipleAmnt4 = reader.GetDouble("PrincipleAmnt4");
                        carddat.PayAfterJudgAmt4 = reader.GetDouble("PayAfterJudgAmt4");
                        carddat.DeptAmnt4 = reader.GetDouble("DeptAmnt4");
                        carddat.LastPayDate4 = reader.GetString("LastPayDate4") ?? "";

                        carddat.CardNo5 = reader.GetString("CardNo5") ?? "";
                        carddat.JudgmentAmnt5 = reader.GetDouble("JudgmentAmnt5");
                        carddat.PrincipleAmnt5 = reader.GetDouble("PrincipleAmnt5");
                        carddat.PayAfterJudgAmt5 = reader.GetDouble("PayAfterJudgAmt5");
                        carddat.DeptAmnt5 = reader.GetDouble("DeptAmnt5");
                        carddat.LastPayDate5 = reader.GetString("LastPayDate5") ?? "";

                        carddat.CardNo6 = reader.GetString("CardNo6") ?? "";
                        carddat.JudgmentAmnt6 = reader.GetDouble("JudgmentAmnt6");
                        carddat.PrincipleAmnt6 = reader.GetDouble("PrincipleAmnt6");
                        carddat.PayAfterJudgAmt6 = reader.GetDouble("PayAfterJudgAmt6");
                        carddat.DeptAmnt6 = reader.GetDouble("DeptAmnt6");
                        carddat.LastPayDate6 = reader.GetString("LastPayDate6") ?? "";

                        carddat.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                        carddat.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        carddat.CustomerTel = security_.DecryptString(reader.GetString("CustomerTel") ?? "");

                        carddat.LegalStatus = reader.GetString("LegalStatus") ?? "";
                        carddat.BlackNo = reader.GetString("BlackNo") ?? "";
                        carddat.RedNo = reader.GetString("RedNo") ?? "";
                        carddat.JudgeDate = reader.GetString("JudgeDate") ?? "";
                        carddat.CourtName = reader.GetString("CourtName") ?? "";
                        carddat.CollectorName = reader.GetString("CollectorName") ?? "";
                        carddat.CollectorTeam = reader.GetString("CollectorTeam") ?? "";
                        carddat.CollectorTel = reader.GetString("CollectorTel") ?? "";
                        carddat.LegalExecDate = reader.GetString("LegalExecDate") ?? "";
                        carddat.LegalExecRemark = reader.GetString("LegalExecRemark") ?? "";
                        carddat.CustomFlag = reader.GetString("CustomFlag") ?? "N";

                        cardall.Add(carddat);
                    }
                }
                connection.Close();
            }
            return cardall;
        }
        public List<DataCPSMaster> doGetDataCPSWithRangeLEDNumber(string startled,string endled) //New
        {
            List<DataCPSMaster> datamasterList = new List<DataCPSMaster>();
            if (connection != null)
            {
                connection.Open();
                var sqlcmd = string.Format(@"SELECT 
                                 ifnull(WorkNo,'') as WorkNo
                                ,ifnull(CaseID,'') as CaseID
                                ,ifnull(CardStatus,'') as CardStatus
                                ,ifnull(LedNumber,'') as LedNumber
                                ,ifnull(CardNo,'') as CardNo
                                ,ifnull(JudgmentAmnt,0) as JudgmentAmnt
                                ,ifnull(PrincipleAmnt,0) as PrincipleAmnt
                                ,ifnull(PayAfterJudgAmt,0) as PayAfterJudgAmt
                                ,ifnull(DeptAmnt,0) as DeptAmnt
                                ,ifnull(LastPayDate,'') as LastPayDate                               
                                ,ifnull(CustomerName,'') as CustomerName
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(CustomerTel,'') as CustomerTel
                                ,ifnull(LegalStatus,'') as LegalStatus
                                ,ifnull(BlackNo,'') as BlackNo
                                ,ifnull(RedNo,'') as RedNo
                                ,ifnull(JudgeDate,'') as JudgeDate
                                ,ifnull(CourtName,'') as CourtName
                                ,ifnull(CollectorName,'') as CollectorName
                                ,ifnull(CollectorTeam,'') as CollectorTeam
                                ,ifnull(CollectorTel,'') as CollectorTel
                                ,ifnull(LegalExecDate,'') as LegalExecDate
                                ,ifnull(LegalExecRemark,'') as LegalExecRemark
                                ,ifnull(CustomFlag,'N') as CustomFlag
                               FROM DataCPSMaster
                               WHERE LedNumber BETWEEN '{0}' AND '{1}';", startled, endled);

                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();                
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        DataCPSMaster masterdat = new DataCPSMaster();
                    
                        masterdat.WorkNo = reader.GetString("WorkNo") ?? "";
                        masterdat.CaseID = reader.GetString("CaseID") ?? "";
                        masterdat.CardStatus = reader.GetString("CardStatus") ?? "";
                        masterdat.LedNumber = reader.GetString("LedNumber") ?? "";
                        masterdat.CardNo = reader.GetString("CardNo") ?? "";
                        masterdat.JudgmentAmnt = reader.GetDouble("JudgmentAmnt");
                        masterdat.PrincipleAmnt = reader.GetDouble("PrincipleAmnt");
                        masterdat.PayAfterJudgAmt = reader.GetDouble("PayAfterJudgAmt");
                        masterdat.DeptAmnt = reader.GetDouble("DeptAmnt");
                        masterdat.LastPayDate = reader.GetString("LastPayDate") ?? "";
                        masterdat.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                        masterdat.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        masterdat.CustomerTel = security_.DecryptString(reader.GetString("CustomerTel") ?? "");
                        masterdat.LegalStatus = reader.GetString("LegalStatus") ?? "";
                        masterdat.BlackNo = reader.GetString("BlackNo") ?? "";
                        masterdat.RedNo = reader.GetString("RedNo") ?? "";
                        masterdat.JudgeDate = reader.GetString("JudgeDate") ?? "";
                        masterdat.CourtName = reader.GetString("CourtName") ?? "";
                        masterdat.CollectorName = reader.GetString("CollectorName") ?? "";
                        masterdat.CollectorTeam = reader.GetString("CollectorTeam") ?? "";
                        masterdat.CollectorTel = reader.GetString("CollectorTel") ?? "";
                        masterdat.LegalExecDate = reader.GetString("LegalExecDate") ?? "";
                        masterdat.LegalExecRemark = reader.GetString("LegalExecRemark") ?? "";
                        masterdat.CustomFlag = reader.GetString("CustomFlag") ?? "N";
                        datamasterList.Add(masterdat);
                    }
                }
                connection.Close();
            }
            return datamasterList;
        }
        public List<DataCPSMaster> doGetDataCPSWithRangeWorkNo(string startworkno, string endworkno)//New
        {
            List<DataCPSMaster> datamasterList = new List<DataCPSMaster>();
            if (connection != null)
            {
                connection.Open();
                var sqlcmd = string.Format(@"SELECT 
                                ifnull(WorkNo,'') as WorkNo
                                ,ifnull(CaseID,'') as CaseID
                                ,ifnull(CardStatus,'') as CardStatus
                                ,ifnull(LedNumber,'') as LedNumber
                                ,ifnull(CardNo,'') as CardNo
                                ,ifnull(JudgmentAmnt,0) as JudgmentAmnt
                                ,ifnull(PrincipleAmnt,0) as PrincipleAmnt
                                ,ifnull(PayAfterJudgAmt,0) as PayAfterJudgAmt
                                ,ifnull(DeptAmnt,0) as DeptAmnt
                                ,ifnull(LastPayDate,'') as LastPayDate                               
                                ,ifnull(CustomerName,'') as CustomerName
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(CustomerTel,'') as CustomerTel
                                ,ifnull(LegalStatus,'') as LegalStatus
                                ,ifnull(BlackNo,'') as BlackNo
                                ,ifnull(RedNo,'') as RedNo
                                ,ifnull(JudgeDate,'') as JudgeDate
                                ,ifnull(CourtName,'') as CourtName
                                ,ifnull(CollectorName,'') as CollectorName
                                ,ifnull(CollectorTeam,'') as CollectorTeam
                                ,ifnull(CollectorTel,'') as CollectorTel
                                ,ifnull(LegalExecDate,'') as LegalExecDate
                                ,ifnull(LegalExecRemark,'') as LegalExecRemark
                                ,ifnull(CustomFlag,'N') as CustomFlag
                               FROM DataCPSMaster
                               WHERE WorkNo BETWEEN '{0}' AND '{1}';", startworkno, endworkno);

                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        DataCPSMaster masterdat = new DataCPSMaster();

                        masterdat.WorkNo = reader.GetString("WorkNo") ?? "";
                        masterdat.CaseID = reader.GetString("CaseID") ?? "";
                        masterdat.CardStatus = reader.GetString("CardStatus") ?? "";
                        masterdat.LedNumber = reader.GetString("LedNumber") ?? "";
                        masterdat.CardNo = reader.GetString("CardNo") ?? "";
                        masterdat.JudgmentAmnt = reader.GetDouble("JudgmentAmnt");
                        masterdat.PrincipleAmnt = reader.GetDouble("PrincipleAmnt");
                        masterdat.PayAfterJudgAmt = reader.GetDouble("PayAfterJudgAmt");
                        masterdat.DeptAmnt = reader.GetDouble("DeptAmnt");
                        masterdat.LastPayDate = reader.GetString("LastPayDate") ?? "";
                        masterdat.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                        masterdat.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        masterdat.CustomerTel = security_.DecryptString(reader.GetString("CustomerTel") ?? "");
                        masterdat.LegalStatus = reader.GetString("LegalStatus") ?? "";
                        masterdat.BlackNo = reader.GetString("BlackNo") ?? "";
                        masterdat.RedNo = reader.GetString("RedNo") ?? "";
                        masterdat.JudgeDate = reader.GetString("JudgeDate") ?? "";
                        masterdat.CourtName = reader.GetString("CourtName") ?? "";
                        masterdat.CollectorName = reader.GetString("CollectorName") ?? "";
                        masterdat.CollectorTeam = reader.GetString("CollectorTeam") ?? "";
                        masterdat.CollectorTel = reader.GetString("CollectorTel") ?? "";
                        masterdat.LegalExecDate = reader.GetString("LegalExecDate") ?? "";
                        masterdat.LegalExecRemark = reader.GetString("LegalExecRemark") ?? "";
                        masterdat.CustomFlag = reader.GetString("CustomFlag") ?? "N";
                        datamasterList.Add(masterdat);
                    }
                }
                connection.Close();

            }
            return datamasterList;
        }
        private void doDeleteCPSMaterData()
        {
            string str_cmd = @"DELETE FROM CPSMaterData;";
            string str_seq = @"DELETE FROM sqlite_sequence WHERE name ='CPSMaterData'";
            if (connection != null)
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand(str_cmd, connection);
                command.ExecuteNonQuery();
                command = new SQLiteCommand(str_seq, connection);
                command.ExecuteNonQuery();
                connection.Close();
            }
        }
        private void doDeleteDataCPSMaster()
        {
            string str_cmd = @"DELETE FROM DataCPSMaster;";
            string str_seq = @"DELETE FROM sqlite_sequence WHERE name ='DataCPSMaster'";
            if (connection != null)
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand(str_cmd, connection);
                command.ExecuteNonQuery();
                command = new SQLiteCommand(str_seq, connection);
                command.ExecuteNonQuery();
                connection.Close();
            }
        }
        private void doDeleteCPSFestData()
        {
            string str_cmd = @"DELETE FROM CPSFestData;";
            string str_seq = @"DELETE FROM sqlite_sequence WHERE name ='CPSFestData'";
            if (connection != null)
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand(str_cmd, connection);
                command.ExecuteNonQuery();
                command = new SQLiteCommand(str_seq, connection);
                command.ExecuteNonQuery();
                connection.Close();
            }
        }
        private void doClearMasterWithFestDataClear() //New
        {
            string cmd_update = @"UPDATE DataCPSMaster
                                SET WorkNo = '',LedNumber = '',LegalExecRemark = ''
                                FROM CPSFestData
                                WHERE (DataCPSMaster.CustomerID = CPSFestData.CustomerID)";
            if (connection != null)
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand(cmd_update, connection);
                int reult = command.ExecuteNonQuery(); 
                connection.Close();
            }
        }
        private void doDeleteCPSCustomData()
        {
            string str_cmd = @"DELETE FROM CPSPayAmnt;";
            string str_seq = @"DELETE FROM sqlite_sequence WHERE name ='CPSPayAmnt'";
            if (connection != null)
            {
                connection.Open();
                SQLiteCommand command = new SQLiteCommand(str_cmd, connection);
                command.ExecuteNonQuery();
                command = new SQLiteCommand(str_seq, connection);
                command.ExecuteNonQuery();
                connection.Close();
            }
        }

        #endregion
        #region DataCPSMaster
        public void doInsertDataCPSMaster(DataTable dataraw, ComboBox[] cmbcolctrl, bool clearmaster, ProgressBar progressLodedata)
        {
            if (dataraw != null)
            {
                progressLodedata.Maximum = dataraw.Rows.Count;
                progressLodedata.Minimum = 0;
                progressLodedata.Step = 1;
                progressLodedata.Visible = true;
                progressLodedata.Value = 0;
                dtResult = dtService.doCreateResultDataTable();
                if (connection != null)
                {
                    if (clearmaster)
                    {
                        doDeleteDataCPSMaster();
                        doDeleteCPSCustomData();
                        doDeleteCPSFestData();
                    }

                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = @"INSERT INTO DataCPSMaster (CaseID,CardStatus,ListNo,CardNo, JudgmentAmnt, PrincipleAmnt, PayAfterJudgAmt, DeptAmnt, LastPayDate, CustomerName, CustomerID, CustomerTel, LegalStatus, BlackNo, RedNo, JudgeDate, CourtName, LegalExecRemark, LegalExecDate, CollectorName, CollectorTeam, CollectorTel)
                                 VALUES($CaseID,$CardStatus,$ListNo,$CardNo, $JudgmentAmnt, $PrincipleAmnt, $PayAfterJudgAmt, $DeptAmnt, $LastPayDate,$CustomerName, $CustomerID,$CustomerTel, $LegalStatus, $BlackNo, $RedNo, $JudgeDate, $CourtName, $LegalExecRemark, $LegalExecDate, $CollectorName, $CollectorTeam, $CollectorTel);";
                        #region Parameter SET
                        var param_CaseID = cmd.CreateParameter();
                        param_CaseID.ParameterName = "$CaseID";
                        cmd.Parameters.Add(param_CaseID);

                        var param_CardStatus = cmd.CreateParameter();
                        param_CardStatus.ParameterName = "$CardStatus";
                        cmd.Parameters.Add(param_CardStatus);

                        var param_ListNo = cmd.CreateParameter();
                        param_ListNo.ParameterName = "$ListNo";
                        cmd.Parameters.Add(param_ListNo);

                        var param_CardNo = cmd.CreateParameter();
                        param_CardNo.ParameterName = "$CardNo";
                        cmd.Parameters.Add(param_CardNo);
                        var param_JudgmentAmnt = cmd.CreateParameter();
                        param_JudgmentAmnt.ParameterName = "$JudgmentAmnt";
                        cmd.Parameters.Add(param_JudgmentAmnt);
                        var param_PrincipleAmnt = cmd.CreateParameter();
                        param_PrincipleAmnt.ParameterName = "$PrincipleAmnt";
                        cmd.Parameters.Add(param_PrincipleAmnt);
                        var param_PayAfterJudgAmt = cmd.CreateParameter();
                        param_PayAfterJudgAmt.ParameterName = "$PayAfterJudgAmt";
                        cmd.Parameters.Add(param_PayAfterJudgAmt);
                        var param_DeptAmnt = cmd.CreateParameter();
                        param_DeptAmnt.ParameterName = "$DeptAmnt";
                        cmd.Parameters.Add(param_DeptAmnt);
                        var param_LastPayDate = cmd.CreateParameter();
                        param_LastPayDate.ParameterName = "$LastPayDate";
                        cmd.Parameters.Add(param_LastPayDate);

                        var param_CustomerName = cmd.CreateParameter();
                        param_CustomerName.ParameterName = "$CustomerName";
                        cmd.Parameters.Add(param_CustomerName);

                        var param_CustomerID = cmd.CreateParameter();
                        param_CustomerID.ParameterName = "$CustomerID";
                        cmd.Parameters.Add(param_CustomerID);

                        var param_CustomerTel = cmd.CreateParameter();
                        param_CustomerTel.ParameterName = "$CustomerTel";
                        cmd.Parameters.Add(param_CustomerTel);

                        var param_LegalStatus = cmd.CreateParameter();
                        param_LegalStatus.ParameterName = "$LegalStatus";
                        cmd.Parameters.Add(param_LegalStatus);

                        var param_BlackNo = cmd.CreateParameter();
                        param_BlackNo.ParameterName = "$BlackNo";
                        cmd.Parameters.Add(param_BlackNo);

                        var param_RedNo = cmd.CreateParameter();
                        param_RedNo.ParameterName = "$RedNo";
                        cmd.Parameters.Add(param_RedNo);

                        var param_JudgeDate = cmd.CreateParameter();
                        param_JudgeDate.ParameterName = "$JudgeDate";
                        cmd.Parameters.Add(param_JudgeDate);

                        var param_CourtName = cmd.CreateParameter();
                        param_CourtName.ParameterName = "$CourtName";
                        cmd.Parameters.Add(param_CourtName);

                        var param_LegalExecRemark = cmd.CreateParameter();
                        param_LegalExecRemark.ParameterName = "$LegalExecRemark";
                        cmd.Parameters.Add(param_LegalExecRemark);

                        var param_LegalExecDate = cmd.CreateParameter();
                        param_LegalExecDate.ParameterName = "$LegalExecDate";
                        cmd.Parameters.Add(param_LegalExecDate);

                        var param_CollectorName = cmd.CreateParameter();
                        param_CollectorName.ParameterName = "$CollectorName";
                        cmd.Parameters.Add(param_CollectorName);

                        var param_CollectorTeam = cmd.CreateParameter();
                        param_CollectorTeam.ParameterName = "$CollectorTeam";
                        cmd.Parameters.Add(param_CollectorTeam);

                        var param_CollectorTel = cmd.CreateParameter();
                        param_CollectorTel.ParameterName = "$CollectorTel";
                        cmd.Parameters.Add(param_CollectorTel);
                        #endregion

                        int listno = 1;
                        string? customerid_befor = string.Empty;
                        string? case_id_brfore = string.Empty;
                        for (int i = 0; i < dataraw.Rows.Count; i++)
                        {
                            param_CardNo.Value = dataraw.Rows[i][cmbcolctrl[0].Text];
                            param_JudgmentAmnt.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[1].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[1].Text]);
                            param_PrincipleAmnt.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[2].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[2].Text]);
                            param_PayAfterJudgAmt.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[3].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[3].Text]);
                            param_DeptAmnt.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[4].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[4].Text]);

                            string? customerid = Convert.ToString(dataraw.Rows[i][cmbcolctrl[7].Text] is null ? string.Empty : dataraw.Rows[i][cmbcolctrl[7].Text]);                            
                            string? customername = Convert.ToString(dataraw.Rows[i][cmbcolctrl[6].Text] is null ? string.Empty : dataraw.Rows[i][cmbcolctrl[6].Text]);
                            string? customertel = Convert.ToString(dataraw.Rows[i][cmbcolctrl[8].Text] is null ? string.Empty : dataraw.Rows[i][cmbcolctrl[8].Text]);
                            string? caseid = Convert.ToString(dataraw.Rows[i][cmbcolctrl[19].Text] is null ? string.Empty : dataraw.Rows[i][cmbcolctrl[19].Text]);
                            if (string.IsNullOrEmpty(customerid_befor))
                            {
                                listno = 1;
                            }
                            else
                            {
                                if (customerid_befor == customerid && case_id_brfore == caseid)
                                {
                                    listno = listno + 1;
                                }
                                else
                                {
                                    listno = 1;
                                }
                            }
                            customerid_befor = customerid;
                            case_id_brfore = caseid;

                            param_ListNo.Value = listno;
                            param_CustomerName.Value = security_.EncryptString(customername);
                            param_CustomerID.Value = security_.EncryptString(customerid);
                            param_CustomerTel.Value = security_.EncryptString(customertel);

                            param_LegalStatus.Value = dataraw.Rows[i][cmbcolctrl[9].Text];
                            param_BlackNo.Value = dataraw.Rows[i][cmbcolctrl[10].Text];
                            param_RedNo.Value = dataraw.Rows[i][cmbcolctrl[11].Text];

                            param_CourtName.Value = dataraw.Rows[i][cmbcolctrl[13].Text];
                            param_CollectorName.Value = dataraw.Rows[i][cmbcolctrl[14].Text];
                            param_CollectorTeam.Value = dataraw.Rows[i][cmbcolctrl[15].Text];
                            param_CollectorTel.Value = dataraw.Rows[i][cmbcolctrl[16].Text];
                            param_LegalExecRemark.Value = dataraw.Rows[i][cmbcolctrl[18].Text];

                            string str_lastpay = datehelper.ConverDateTODBStr(dataraw.Rows[i], cmbcolctrl[5].Text);
                            string str_judedate = datehelper.ConverDateTODBStr(dataraw.Rows[i], cmbcolctrl[12].Text);
                            string str_execdate = datehelper.ConverDateTODBStr(dataraw.Rows[i], cmbcolctrl[17].Text);
                            param_LastPayDate.Value = str_lastpay;
                            param_JudgeDate.Value = str_judedate;
                            param_LegalExecDate.Value = str_execdate;

                            param_CaseID.Value = caseid;
                            param_CardStatus.Value = dataraw.Rows[i][cmbcolctrl[20].Text];

                            int result_i = cmd.ExecuteNonQuery();

                            if (result_i > 0)
                            {
                                if (!string.IsNullOrEmpty(customerid)) doSetResultData(ref dtResult, customerid, "Y", "");
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(customerid)) doSetResultData(ref dtResult, customerid, "N", "");
                            }
                            if (str_lastpay.Contains("10000101")) if (!string.IsNullOrEmpty(customerid)) doSetResultData(ref dtResult, customerid, "N", "วันที่ผิด LastPay");
                            if (str_judedate.Contains("10000101")) if (!string.IsNullOrEmpty(customerid)) doSetResultData(ref dtResult, customerid, "N", "วันที่ผิด Jude");
                            if (str_execdate.Contains("10000101")) if (!string.IsNullOrEmpty(customerid)) doSetResultData(ref dtResult, customerid, "N", "วันที่ผิด Exec");
                            progressLodedata.Value = i + 1;
                        }
                        transaction.Commit();
                    }
                    connection.Close();
                }
                progressLodedata.Visible = false;
            }
        }
        public List<DataCPSPerson> doGetDataCPSMaterWithCustomerID(string customerid)
        {
            List<DataCPSPerson> cardall = new List<DataCPSPerson>();
            if (string.IsNullOrEmpty(customerid)) return cardall;

            string customerid_where = security_.EncryptString(customerid);
            var sqlcmd = string.Format(@"SELECT 
                                 ifnull(CaseID,'') as CaseID
                                ,ifnull(CardStatus,'') as CardStatus
                                ,ifnull(ListNo,0) as ListNo
                                ,ifnull(WorkNo,'') as WorkNo
                                ,ifnull(LedNumber,'') as LedNumber
                                ,ifnull(CardNo,'') as CardNo
                                ,ifnull(JudgmentAmnt,0) as JudgmentAmnt
                                ,ifnull(PrincipleAmnt,0) as PrincipleAmnt
                                ,ifnull(PayAfterJudgAmt,0) as PayAfterJudgAmt
                                ,ifnull(DeptAmnt,0) as DeptAmnt
                                ,ifnull(LastPayDate,'') as LastPayDate                                
                                ,ifnull(CustomerName,'') as CustomerName
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(CustomerTel,'') as CustomerTel
                                ,ifnull(LegalStatus,'') as LegalStatus
                                ,ifnull(BlackNo,'') as BlackNo
                                ,ifnull(RedNo,'') as RedNo
                                ,ifnull(JudgeDate,'') as JudgeDate
                                ,ifnull(CourtName,'') as CourtName
                                ,ifnull(CollectorName,'') as CollectorName
                                ,ifnull(CollectorTeam,'') as CollectorTeam
                                ,ifnull(CollectorTel,'') as CollectorTel
                                ,ifnull(LegalExecDate,'') as LegalExecDate
                                ,ifnull(LegalExecRemark,'') as LegalExecRemark
                                ,ifnull(CustomFlag,'N') as CustomFlag
                               FROM DataCPSMaster Order by CustomerID");
            if (connection != null)
            {
                connection.Open();
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();
                string custid_old = string.Empty;
                if (reader.HasRows)
                {
                    DataCPSPerson carddat = new DataCPSPerson();
                    while (reader.Read())
                    {
                        string cust_id = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        if (!string.IsNullOrEmpty(cust_id))
                        {
                            carddat.CaseID =reader.GetString("CaseID") ?? "";
                            carddat.CardStatus =reader.GetString("CardStatus") ?? "";
                            carddat.WorkNo = reader.GetString("WorkNo") ?? "";
                            carddat.LedNumber = reader.GetString("LedNumber") ?? "";
                            carddat.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                            carddat.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                            carddat.CustomerTel = security_.DecryptString(reader.GetString("CustomerTel") ?? "");
                            carddat.LegalStatus = reader.GetString("LegalStatus") ?? "";
                            carddat.BlackNo = reader.GetString("BlackNo") ?? "";
                            carddat.RedNo = reader.GetString("RedNo") ?? "";
                            carddat.JudgeDate = reader.GetString("JudgeDate") ?? "";
                            carddat.CourtName = reader.GetString("CourtName") ?? "";
                            carddat.CollectorName = reader.GetString("CollectorName") ?? "";
                            carddat.CollectorTeam = reader.GetString("CollectorTeam") ?? "";
                            carddat.CollectorTel = reader.GetString("CollectorTel") ?? "";
                            carddat.LegalExecDate = reader.GetString("LegalExecDate") ?? "";
                            carddat.LegalExecRemark = reader.GetString("LegalExecRemark") ?? "";
                            carddat.CustomFlag = reader.GetString("CustomFlag") ?? "N";
                            int listno = reader.GetInt32("ListNo");
                            doSetDataCardByListno(reader, listno, ref carddat);                           
                        }
                    }                   
                    cardall.Add(carddat);
                }

                connection.Close() ;
            }
            return cardall;
        }
        public List<DataCPSPerson> doGetDataCPSMasterGroupID()
        {
            List<DataCPSPerson> cardall = new List<DataCPSPerson>();
            var sqlcmd = string.Format(@"SELECT  ifnull(CustomerID,'') as CustomerID FROM DataCPSMaster Group By CustomerID;");
            if (connection != null)
            {
                connection.Open();
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();

                if (reader.HasRows)
                {                   
                    while (reader.Read())
                    {
                        DataCPSPerson carddat = new DataCPSPerson();
                        carddat.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        cardall.Add(carddat);
                    }                    
                }
                connection.Close();
            }
            return cardall;
        }
        public List<DataCPSMaster> doGetDataCPSMasterAll()
        {
            List<DataCPSMaster> masterall = new List<DataCPSMaster>();
            if (connection != null)
            {
                connection.Open();
                var sqlcmd = string.Format(@"SELECT 
                                 ifnull(CaseID,'') as CaseID
                                ,ifnull(CardStatus,'') as CardStatus
                                ,ifnull(ListNo,0) as ListNo
                                ,ifnull(WorkNo,'') as WorkNo
                                ,ifnull(LedNumber,'') as LedNumber
                                ,ifnull(CardNo,'') as CardNo
                                ,ifnull(JudgmentAmnt,0) as JudgmentAmnt
                                ,ifnull(PrincipleAmnt,0) as PrincipleAmnt
                                ,ifnull(PayAfterJudgAmt,0) as PayAfterJudgAmt
                                ,ifnull(DeptAmnt,0) as DeptAmnt
                                ,ifnull(LastPayDate,'') as LastPayDate                                
                                ,ifnull(CustomerName,'') as CustomerName
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(CustomerTel,'') as CustomerTel
                                ,ifnull(LegalStatus,'') as LegalStatus
                                ,ifnull(BlackNo,'') as BlackNo
                                ,ifnull(RedNo,'') as RedNo
                                ,ifnull(JudgeDate,'') as JudgeDate
                                ,ifnull(CourtName,'') as CourtName
                                ,ifnull(CollectorName,'') as CollectorName
                                ,ifnull(CollectorTeam,'') as CollectorTeam
                                ,ifnull(CollectorTel,'') as CollectorTel
                                ,ifnull(LegalExecDate,'') as LegalExecDate
                                ,ifnull(LegalExecRemark,'') as LegalExecRemark
                                ,ifnull(CustomFlag,'N') as CustomFlag
                               FROM DataCPSMaster
                               Order By CustomerID ASC");
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        DataCPSMaster carddat = new DataCPSMaster();
                        carddat.CaseID = reader.GetString("CaseID") ?? "";
                        carddat.CardStatus = reader.GetString("CardStatus") ?? "";
                        carddat.WorkNo = reader.GetString("WorkNo") ?? "";
                        carddat.LedNumber = reader.GetString("LedNumber") ?? "";
                        carddat.CardNo = reader.GetString("CardNo") ?? "";
                        carddat.JudgmentAmnt = reader.GetDouble("JudgmentAmnt");
                        carddat.PrincipleAmnt = reader.GetDouble("PrincipleAmnt");
                        carddat.PayAfterJudgAmt = reader.GetDouble("PayAfterJudgAmt");
                        carddat.DeptAmnt = reader.GetDouble("DeptAmnt");
                        carddat.LastPayDate = reader.GetString("LastPayDate") ?? "";

                        carddat.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                        carddat.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        carddat.CustomerTel = security_.DecryptString(reader.GetString("CustomerTel") ?? "");

                        carddat.LegalStatus = reader.GetString("LegalStatus") ?? "";
                        carddat.BlackNo = reader.GetString("BlackNo") ?? "";
                        carddat.RedNo = reader.GetString("RedNo") ?? "";
                        carddat.JudgeDate = reader.GetString("JudgeDate") ?? "";
                        carddat.CourtName = reader.GetString("CourtName") ?? "";
                        carddat.CollectorName = reader.GetString("CollectorName") ?? "";
                        carddat.CollectorTeam = reader.GetString("CollectorTeam") ?? "";
                        carddat.CollectorTel = reader.GetString("CollectorTel") ?? "";
                        carddat.LegalExecDate = reader.GetString("LegalExecDate") ?? "";
                        carddat.LegalExecRemark = reader.GetString("LegalExecRemark") ?? "";
                        carddat.CustomFlag = reader.GetString("CustomFlag") ?? "N";

                        masterall.Add(carddat);
                    }
                }
                connection.Close();
            }
            return masterall;
        }
        private void doSetDataCardByListno(SQLiteDataReader reder,int listno,ref DataCPSPerson carddt)
        {
            string _CardNo = reder.GetString("CardNo");
            double _JudgmentAmnt = reder.GetDouble("JudgmentAmnt");
            double _PrincipleAmnt = reder.GetDouble("PrincipleAmnt");
            double _PayAfterJudgAmnt = reder.GetDouble("PayAfterJudgAmt");
            double _DeptAmnt = reder.GetDouble("DeptAmnt");
            string _LastPayDate = reder.GetString("LastPayDate");
            switch (listno)
            {
                case 1:
                    carddt.CardNo1 = _CardNo;
                    carddt.JudgmentAmnt1 = _JudgmentAmnt;
                    carddt.PrincipleAmnt1 = _PrincipleAmnt;
                    carddt.PayAfterJudgAmt1 = _PayAfterJudgAmnt;
                    carddt.DeptAmnt1 = _DeptAmnt;
                    carddt.LastPayDate1 = _LastPayDate;
                    break;
                case 2:
                    carddt.CardNo2 = _CardNo;
                    carddt.JudgmentAmnt2 = _JudgmentAmnt;
                    carddt.PrincipleAmnt2 = _PrincipleAmnt;
                    carddt.PayAfterJudgAmt2 = _PayAfterJudgAmnt;
                    carddt.DeptAmnt2 = _DeptAmnt;
                    carddt.LastPayDate2 = _LastPayDate;
                    break;
                case 3:
                    carddt.CardNo3 = _CardNo;
                    carddt.JudgmentAmnt3 = _JudgmentAmnt;
                    carddt.PrincipleAmnt3 = _PrincipleAmnt;
                    carddt.PayAfterJudgAmt3 = _PayAfterJudgAmnt;
                    carddt.DeptAmnt3 = _DeptAmnt;
                    carddt.LastPayDate3 = _LastPayDate;
                    break;
                case 4:
                    carddt.CardNo4 = _CardNo;
                    carddt.JudgmentAmnt4 = _JudgmentAmnt;
                    carddt.PrincipleAmnt4 = _PrincipleAmnt;
                    carddt.PayAfterJudgAmt4 = _PayAfterJudgAmnt;
                    carddt.DeptAmnt4 = _DeptAmnt;
                    carddt.LastPayDate4 = _LastPayDate;
                    break;
                case 5:
                    carddt.CardNo5 = _CardNo;
                    carddt.JudgmentAmnt5 = _JudgmentAmnt;
                    carddt.PrincipleAmnt5 = _PrincipleAmnt;
                    carddt.PayAfterJudgAmt5 = _PayAfterJudgAmnt;
                    carddt.DeptAmnt5 = _DeptAmnt;
                    carddt.LastPayDate5 = _LastPayDate;
                    break;
                case 6:
                    carddt.CardNo6 = _CardNo;
                    carddt.JudgmentAmnt6 = _JudgmentAmnt;
                    carddt.PrincipleAmnt6 = _PrincipleAmnt;
                    carddt.PayAfterJudgAmt6 = _PayAfterJudgAmnt;
                    carddt.DeptAmnt6 = _DeptAmnt;
                    carddt.LastPayDate6 = _LastPayDate;
                    break;
            }
        }
        #endregion
        #region CustomCPS
        public void doInsertCustomData(DataTable dataraw, ComboBox[] cmbcolctrl, ProgressBar progressLodedata, bool cleardata)
        {
            if (dataraw != null)
            {
                progressLodedata.Maximum = dataraw.Rows.Count;
                progressLodedata.Minimum = 0;
                progressLodedata.Step = 1;
                progressLodedata.Value = 0;
                progressLodedata.Visible = true;
                if (connection != null)
                {
                    if (cleardata)
                    {
                        doClearMasterWitCustomClear();
                        doDeleteCPSCustomData();
                    }
                    connection.Open();
                    dtResult = dtService.doCreateResultDataTable();
                    using (var transaction = connection.BeginTransaction())
                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = @"INSERT INTO CPSPayAmnt (CaseID,CustomerID,CustomerName,WorkNo,LedNumber,CardNo1,AccCloseAmnt1,AccClose6Amnt1,AccClose12Amnt1,AccClose24Amnt1
                                           ,CardNo2,AccCloseAmnt2,AccClose6Amnt2,AccClose12Amnt2,AccClose24Amnt2,CardNo3,AccCloseAmnt3,AccClose6Amnt3,AccClose12Amnt3,AccClose24Amnt3,CardNo4,AccCloseAmnt4
                                           ,AccClose6Amnt4,AccClose12Amnt4,AccClose24Amnt4,CardNo5,AccCloseAmnt5,AccClose6Amnt5,AccClose12Amnt5,AccClose24Amnt5,CardNo6,AccCloseAmnt6,AccClose6Amnt6,AccClose12Amnt6
                                           ,Installment6Amnt1,Installment6Amnt2,Installment6Amnt3,Installment6Amnt4,Installment6Amnt5,Installment6Amnt6,Installment12Amnt1,Installment12Amnt2,Installment12Amnt3
                                           ,Installment12Amnt4,Installment12Amnt5,Installment12Amnt6,Installment24Amnt1,Installment24Amnt2,Installment24Amnt3,Installment24Amnt4,Installment24Amnt5,Installment24Amnt6
                                           ,AccClose24Amnt6,LegalExecRemark) VALUES($CaseID,$CustomerID,$CustomerName,$WorkNo,$LedNumber,$CardNo1,$AccCloseAmnt1,$AccClose6Amnt1,$AccClose12Amnt1,$AccClose24Amnt1,$CardNo2,$AccCloseAmnt2
                                           ,$AccClose6Amnt2,$AccClose12Amnt2,$AccClose24Amnt2,$CardNo3,$AccCloseAmnt3,$AccClose6Amnt3,$AccClose12Amnt3,$AccClose24Amnt3,$CardNo4,$AccCloseAmnt4,$AccClose6Amnt4
                                           ,$AccClose12Amnt4,$AccClose24Amnt4,$CardNo5,$AccCloseAmnt5,$AccClose6Amnt5,$AccClose12Amnt5,$AccClose24Amnt5,$CardNo6,$AccCloseAmnt6,$AccClose6Amnt6,$AccClose12Amnt6
                                           ,$Installment6Amnt1,$Installment6Amnt2,$Installment6Amnt3,$Installment6Amnt4,$Installment6Amnt5,$Installment6Amnt6,$Installment12Amnt1,$Installment12Amnt2,$Installment12Amnt3
                                           ,$Installment12Amnt4,$Installment12Amnt5,$Installment12Amnt6,$Installment24Amnt1,$Installment24Amnt2,$Installment24Amnt3,$Installment24Amnt4,$Installment24Amnt5,$Installment24Amnt6
                                           ,$AccClose24Amnt6,$LegalExecRemark);";

                        #region Parameter SET
                        var param_CaseID = cmd.CreateParameter();
                        param_CaseID.ParameterName = "$CaseID";
                        cmd.Parameters.Add(param_CaseID);

                        var param_CustomerID = cmd.CreateParameter();
                        param_CustomerID.ParameterName = "$CustomerID";
                        cmd.Parameters.Add(param_CustomerID);

                        var param_CustomerName = cmd.CreateParameter();
                        param_CustomerName.ParameterName = "$CustomerName";
                        cmd.Parameters.Add(param_CustomerName);

                        var param_LedNumber = cmd.CreateParameter();
                        param_LedNumber.ParameterName = "$LedNumber";
                        cmd.Parameters.Add(param_LedNumber);

                        var param_WorkNo = cmd.CreateParameter();
                        param_WorkNo.ParameterName = "$WorkNo";
                        cmd.Parameters.Add(param_WorkNo);

                        var param_CardNo1 = cmd.CreateParameter();
                        param_CardNo1.ParameterName = "$CardNo1";
                        cmd.Parameters.Add(param_CardNo1);

                        var param_AccCloseAmnt1 = cmd.CreateParameter();
                        param_AccCloseAmnt1.ParameterName = "$AccCloseAmnt1";
                        cmd.Parameters.Add(param_AccCloseAmnt1);

                        var param_AccClose6Amnt1 = cmd.CreateParameter();
                        param_AccClose6Amnt1.ParameterName = "$AccClose6Amnt1";
                        cmd.Parameters.Add(param_AccClose6Amnt1);

                        var param_AccClose12Amnt1 = cmd.CreateParameter();
                        param_AccClose12Amnt1.ParameterName = "$AccClose12Amnt1";
                        cmd.Parameters.Add(param_AccClose12Amnt1);

                        var param_AccClose24Amnt1 = cmd.CreateParameter();
                        param_AccClose24Amnt1.ParameterName = "$AccClose24Amnt1";
                        cmd.Parameters.Add(param_AccClose24Amnt1);

                        var param_CardNo2 = cmd.CreateParameter();
                        param_CardNo2.ParameterName = "$CardNo2";
                        cmd.Parameters.Add(param_CardNo2);

                        var param_AccCloseAmnt2 = cmd.CreateParameter();
                        param_AccCloseAmnt2.ParameterName = "$AccCloseAmnt2";
                        cmd.Parameters.Add(param_AccCloseAmnt2);

                        var param_AccClose6Amnt2 = cmd.CreateParameter();
                        param_AccClose6Amnt2.ParameterName = "$AccClose6Amnt2";
                        cmd.Parameters.Add(param_AccClose6Amnt2);

                        var param_AccClose12Amnt2 = cmd.CreateParameter();
                        param_AccClose12Amnt2.ParameterName = "$AccClose12Amnt2";
                        cmd.Parameters.Add(param_AccClose12Amnt2);

                        var param_AccClose24Amnt2 = cmd.CreateParameter();
                        param_AccClose24Amnt2.ParameterName = "$AccClose24Amnt2";
                        cmd.Parameters.Add(param_AccClose24Amnt2);

                        var param_CardNo3 = cmd.CreateParameter();
                        param_CardNo3.ParameterName = "$CardNo3";
                        cmd.Parameters.Add(param_CardNo3);

                        var param_AccCloseAmnt3 = cmd.CreateParameter();
                        param_AccCloseAmnt3.ParameterName = "$AccCloseAmnt3";
                        cmd.Parameters.Add(param_AccCloseAmnt3);

                        var param_AccClose6Amnt3 = cmd.CreateParameter();
                        param_AccClose6Amnt3.ParameterName = "$AccClose6Amnt3";
                        cmd.Parameters.Add(param_AccClose6Amnt3);

                        var param_AccClose12Amnt3 = cmd.CreateParameter();
                        param_AccClose12Amnt3.ParameterName = "$AccClose12Amnt3";
                        cmd.Parameters.Add(param_AccClose12Amnt3);

                        var param_AccClose24Amnt3 = cmd.CreateParameter();
                        param_AccClose24Amnt3.ParameterName = "$AccClose24Amnt3";
                        cmd.Parameters.Add(param_AccClose24Amnt3);

                        var param_CardNo4 = cmd.CreateParameter();
                        param_CardNo4.ParameterName = "$CardNo4";
                        cmd.Parameters.Add(param_CardNo4);

                        var param_AccCloseAmnt4 = cmd.CreateParameter();
                        param_AccCloseAmnt4.ParameterName = "$AccCloseAmnt4";
                        cmd.Parameters.Add(param_AccCloseAmnt4);

                        var param_AccClose6Amnt4 = cmd.CreateParameter();
                        param_AccClose6Amnt4.ParameterName = "$AccClose6Amnt4";
                        cmd.Parameters.Add(param_AccClose6Amnt4);

                        var param_AccClose12Amnt4 = cmd.CreateParameter();
                        param_AccClose12Amnt4.ParameterName = "$AccClose12Amnt4";
                        cmd.Parameters.Add(param_AccClose12Amnt4);

                        var param_AccClose24Amnt4 = cmd.CreateParameter();
                        param_AccClose24Amnt4.ParameterName = "$AccClose24Amnt4";
                        cmd.Parameters.Add(param_AccClose24Amnt4);

                        var param_CardNo5 = cmd.CreateParameter();
                        param_CardNo5.ParameterName = "$CardNo5";
                        cmd.Parameters.Add(param_CardNo5);

                        var param_AccCloseAmnt5 = cmd.CreateParameter();
                        param_AccCloseAmnt5.ParameterName = "$AccCloseAmnt5";
                        cmd.Parameters.Add(param_AccCloseAmnt5);

                        var param_AccClose6Amnt5 = cmd.CreateParameter();
                        param_AccClose6Amnt5.ParameterName = "$AccClose6Amnt5";
                        cmd.Parameters.Add(param_AccClose6Amnt5);

                        var param_AccClose12Amnt5 = cmd.CreateParameter();
                        param_AccClose12Amnt5.ParameterName = "$AccClose12Amnt5";
                        cmd.Parameters.Add(param_AccClose12Amnt5);

                        var param_AccClose24Amnt5 = cmd.CreateParameter();
                        param_AccClose24Amnt5.ParameterName = "$AccClose24Amnt5";
                        cmd.Parameters.Add(param_AccClose24Amnt5);

                        var param_CardNo6 = cmd.CreateParameter();
                        param_CardNo6.ParameterName = "$CardNo6";
                        cmd.Parameters.Add(param_CardNo6);

                        var param_AccCloseAmnt6 = cmd.CreateParameter();
                        param_AccCloseAmnt6.ParameterName = "$AccCloseAmnt6";
                        cmd.Parameters.Add(param_AccCloseAmnt6);

                        var param_AccClose6Amnt6 = cmd.CreateParameter();
                        param_AccClose6Amnt6.ParameterName = "$AccClose6Amnt6";
                        cmd.Parameters.Add(param_AccClose6Amnt6);

                        var param_AccClose12Amnt6 = cmd.CreateParameter();
                        param_AccClose12Amnt6.ParameterName = "$AccClose12Amnt6";
                        cmd.Parameters.Add(param_AccClose12Amnt6);

                        var param_AccClose24Amnt6 = cmd.CreateParameter();
                        param_AccClose24Amnt6.ParameterName = "$AccClose24Amnt6";
                        cmd.Parameters.Add(param_AccClose24Amnt6);

                        var param_LegalExecRemark = cmd.CreateParameter();
                        param_LegalExecRemark.ParameterName = "$LegalExecRemark";
                        cmd.Parameters.Add(param_LegalExecRemark);

                        var param_Installment6Amnt1 = cmd.CreateParameter();
                        param_Installment6Amnt1.ParameterName = "$Installment6Amnt1";
                        cmd.Parameters.Add(param_Installment6Amnt1);

                        var param_Installment12Amnt1 = cmd.CreateParameter();
                        param_Installment12Amnt1.ParameterName = "$Installment12Amnt1";
                        cmd.Parameters.Add(param_Installment12Amnt1);

                        var param_Installment24Amnt1 = cmd.CreateParameter();
                        param_Installment24Amnt1.ParameterName = "$Installment24Amnt1";
                        cmd.Parameters.Add(param_Installment24Amnt1);

                        var param_Installment6Amnt2 = cmd.CreateParameter();
                        param_Installment6Amnt2.ParameterName = "$Installment6Amnt2";
                        cmd.Parameters.Add(param_Installment6Amnt2);

                        var param_Installment12Amnt2 = cmd.CreateParameter();
                        param_Installment12Amnt2.ParameterName = "$Installment12Amnt2";
                        cmd.Parameters.Add(param_Installment12Amnt2);

                        var param_Installment24Amnt2 = cmd.CreateParameter();
                        param_Installment24Amnt2.ParameterName = "$Installment24Amnt2";
                        cmd.Parameters.Add(param_Installment24Amnt2);

                        var param_Installment6Amnt3 = cmd.CreateParameter();
                        param_Installment6Amnt3.ParameterName = "$Installment6Amnt3";
                        cmd.Parameters.Add(param_Installment6Amnt3);

                        var param_Installment12Amnt3 = cmd.CreateParameter();
                        param_Installment12Amnt3.ParameterName = "$Installment12Amnt3";
                        cmd.Parameters.Add(param_Installment12Amnt3);

                        var param_Installment24Amnt3 = cmd.CreateParameter();
                        param_Installment24Amnt3.ParameterName = "$Installment24Amnt3";
                        cmd.Parameters.Add(param_Installment24Amnt3);

                        var param_Installment6Amnt4 = cmd.CreateParameter();
                        param_Installment6Amnt4.ParameterName = "$Installment6Amnt4";
                        cmd.Parameters.Add(param_Installment6Amnt4);

                        var param_Installment12Amnt4 = cmd.CreateParameter();
                        param_Installment12Amnt4.ParameterName = "$Installment12Amnt4";
                        cmd.Parameters.Add(param_Installment12Amnt4);

                        var param_Installment24Amnt4 = cmd.CreateParameter();
                        param_Installment24Amnt4.ParameterName = "$Installment24Amnt4";
                        cmd.Parameters.Add(param_Installment24Amnt4);

                        var param_Installment6Amnt5 = cmd.CreateParameter();
                        param_Installment6Amnt5.ParameterName = "$Installment6Amnt5";
                        cmd.Parameters.Add(param_Installment6Amnt5);

                        var param_Installment12Amnt5 = cmd.CreateParameter();
                        param_Installment12Amnt5.ParameterName = "$Installment12Amnt5";
                        cmd.Parameters.Add(param_Installment12Amnt5);

                        var param_Installment24Amnt5 = cmd.CreateParameter();
                        param_Installment24Amnt5.ParameterName = "$Installment24Amnt5";
                        cmd.Parameters.Add(param_Installment24Amnt5);

                        var param_Installment6Amnt6 = cmd.CreateParameter();
                        param_Installment6Amnt6.ParameterName = "$Installment6Amnt6";
                        cmd.Parameters.Add(param_Installment6Amnt6);

                        var param_Installment12Amnt6 = cmd.CreateParameter();
                        param_Installment12Amnt6.ParameterName = "$Installment12Amnt6";
                        cmd.Parameters.Add(param_Installment12Amnt6);

                        var param_Installment24Amnt6 = cmd.CreateParameter();
                        param_Installment24Amnt6.ParameterName = "$Installment24Amnt6";
                        cmd.Parameters.Add(param_Installment24Amnt6);

                        #endregion

                        for (int i = 0; i < dataraw.Rows.Count; i++)
                        {
                            string? customerid = Convert.ToString(dataraw.Rows[i][cmbcolctrl[0].Text] is null?string.Empty: dataraw.Rows[i][cmbcolctrl[0].Text]);
                            string? customertel = Convert.ToString(dataraw.Rows[i][cmbcolctrl[1].Text] is null ? string.Empty : dataraw.Rows[i][cmbcolctrl[1].Text]);
                            param_CustomerID.Value = security_.EncryptString(customerid);
                            param_CustomerName.Value = security_.EncryptString(customertel);
                            param_CaseID.Value = Convert.ToString(dataraw.Rows[i][cmbcolctrl[53].Text]); 
                            param_LedNumber.Value = Convert.ToString(dataraw.Rows[i][cmbcolctrl[2].Text]);
                            param_WorkNo.Value = Convert.ToString(dataraw.Rows[i][cmbcolctrl[3].Text]);                           

                            param_CardNo1.Value = dataraw.Rows[i][cmbcolctrl[4].Text];
                            param_CardNo2.Value = dataraw.Rows[i][cmbcolctrl[5].Text];
                            param_CardNo3.Value = dataraw.Rows[i][cmbcolctrl[6].Text];
                            param_CardNo4.Value = dataraw.Rows[i][cmbcolctrl[7].Text];
                            param_CardNo5.Value = dataraw.Rows[i][cmbcolctrl[8].Text];
                            param_CardNo6.Value = dataraw.Rows[i][cmbcolctrl[9].Text];

                            param_AccCloseAmnt1.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[10].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[10].Text]);
                            param_AccCloseAmnt2.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[11].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[11].Text]);
                            param_AccCloseAmnt3.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[12].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[12].Text]);
                            param_AccCloseAmnt4.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[13].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[13].Text]);
                            param_AccCloseAmnt5.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[14].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[14].Text]);
                            param_AccCloseAmnt6.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[15].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[15].Text]);
                                                                                                                                                                                         
                            param_AccClose6Amnt1.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[16].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[16].Text]);
                            param_AccClose6Amnt2.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[17].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[17].Text]);
                            param_AccClose6Amnt3.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[18].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[18].Text]);
                            param_AccClose6Amnt4.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[19].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[19].Text]);
                            param_AccClose6Amnt5.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[20].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[20].Text]);
                            param_AccClose6Amnt6.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[21].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[21].Text]);

                            param_AccClose12Amnt1.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[22].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[22].Text]);
                            param_AccClose12Amnt2.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[23].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[23].Text]);
                            param_AccClose12Amnt3.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[24].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[24].Text]);
                            param_AccClose12Amnt4.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[25].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[25].Text]);
                            param_AccClose12Amnt5.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[26].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[26].Text]);
                            param_AccClose12Amnt6.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[27].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[27].Text]);

                            param_AccClose24Amnt1.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[28].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[28].Text]);  
                            param_AccClose24Amnt2.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[29].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[29].Text]);
                            param_AccClose24Amnt3.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[30].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[30].Text]);
                            param_AccClose24Amnt4.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[31].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[31].Text]);
                            param_AccClose24Amnt5.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[32].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[32].Text]);
                            param_AccClose24Amnt6.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[33].Text].ToString())? 0 : dataraw.Rows[i][cmbcolctrl[33].Text]);
                            param_LegalExecRemark.Value = Convert.ToString(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[34].Text].ToString())? "" : dataraw.Rows[i][cmbcolctrl[34].Text]);

                            param_Installment6Amnt1.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[35].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[35].Text]);
                            param_Installment6Amnt2.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[36].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[36].Text]);
                            param_Installment6Amnt3.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[37].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[37].Text]);
                            param_Installment6Amnt4.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[38].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[38].Text]);
                            param_Installment6Amnt5.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[39].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[39].Text]);
                            param_Installment6Amnt6.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[40].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[40].Text]);
                            param_Installment12Amnt1.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[41].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[41].Text]);
                            param_Installment12Amnt2.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[42].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[42].Text]);
                            param_Installment12Amnt3.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[43].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[43].Text]);
                            param_Installment12Amnt4.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[44].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[44].Text]);
                            param_Installment12Amnt5.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[45].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[45].Text]);
                            param_Installment12Amnt6.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[46].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[46].Text]);
                            param_Installment24Amnt1.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[47].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[47].Text]);
                            param_Installment24Amnt2.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[48].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[48].Text]);
                            param_Installment24Amnt3.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[49].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[49].Text]);
                            param_Installment24Amnt4.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[50].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[50].Text]);
                            param_Installment24Amnt5.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[51].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[51].Text]);
                            param_Installment24Amnt6.Value = Convert.ToDouble(string.IsNullOrEmpty(dataraw.Rows[i][cmbcolctrl[52].Text].ToString()) ? 0 : dataraw.Rows[i][cmbcolctrl[52].Text]);

                            int result = cmd.ExecuteNonQuery();
                            if(result >= 0)
                            {
                                string? customerid_ = Convert.ToString(param_CustomerID.Value??"");
                                if (result > 0)
                                {

                                    if (!string.IsNullOrEmpty(customerid)) doSetResultData(ref dtResult, customerid, "Y", "");
                                }
                                else
                                {
                                    if (!string.IsNullOrEmpty(customerid)) doSetResultData(ref dtResult, customerid, "N", "");
                                }
                            }
                            progressLodedata.Value = i + 1;
                        }
                        transaction.Commit();
                    }
                    doUpdateCustomFlag();
                    connection.Close();
                   
                }
                progressLodedata.Visible = false;
            }
        }
        public List<FestCustom> doGetDataCustomWithID(string customerid,string caseid)
        {
            List<FestCustom> customdataall = new List<FestCustom>();
            if (string.IsNullOrEmpty(customerid)) return customdataall;
            if (connection != null)
            {
                string customerwhere = security_.EncryptString(customerid);
                connection.Open();
                var sqlcmd = string.Format(@" SELECT
                                             ifnull(WorkNo,'') as WorkNo
                                            ,ifnull(LedNumber,'') as LedNumber
                                            ,ifnull(CustomerID,'') as CustomerID
                                            ,ifnull(CaseID,'') as CaseID
                                            ,ifnull(CustomerName,'') as CustomerName
                                            ,ifnull(CardNo1,'') as CardNo1
                                            ,ifnull(AccCloseAmnt1,0) as AccCloseAmnt1
                                            ,ifnull(AccClose6Amnt1,0) as AccClose6Amnt1
                                            ,ifnull(AccClose12Amnt1,0) as AccClose12Amnt1
                                            ,ifnull(AccClose24Amnt1,0) as AccClose24Amnt1
                                            ,ifnull(CardNo2,'') as CardNo2
                                            ,ifnull(AccCloseAmnt2,0) as AccCloseAmnt2
                                            ,ifnull(AccClose6Amnt2,0) as AccClose6Amnt2
                                            ,ifnull(AccClose12Amnt2,0) as AccClose12Amnt2
                                            ,ifnull(AccClose24Amnt2,0) as AccClose24Amnt2
                                            ,ifnull(CardNo3,'') as CardNo3
                                            ,ifnull(AccCloseAmnt3,0) as AccCloseAmnt3
                                            ,ifnull(AccClose6Amnt3,0) as AccClose6Amnt3
                                            ,ifnull(AccClose12Amnt3,0) as AccClose12Amnt3
                                            ,ifnull(AccClose24Amnt3,0) as AccClose24Amnt3
                                            ,ifnull(CardNo4,'') as CardNo4
                                            ,ifnull(AccCloseAmnt4,0) as AccCloseAmnt4
                                            ,ifnull(AccClose6Amnt4,0) as AccClose6Amnt4
                                            ,ifnull(AccClose12Amnt4,0) as AccClose12Amnt4
                                            ,ifnull(AccClose24Amnt4,0) as AccClose24Amnt4
                                            ,ifnull(CardNo5,'') as CardNo5
                                            ,ifnull(AccCloseAmnt5,0) as AccCloseAmnt5
                                            ,ifnull(AccClose6Amnt5,0) as AccClose6Amnt5
                                            ,ifnull(AccClose12Amnt5,0) as AccClose12Amnt5
                                            ,ifnull(AccClose24Amnt5,0) as AccClose24Amnt5
                                            ,ifnull(CardNo6,'') as CardNo6
                                            ,ifnull(AccCloseAmnt6,0) as AccCloseAmnt6
                                            ,ifnull(AccClose6Amnt6,0) as AccClose6Amnt6
                                            ,ifnull(AccClose12Amnt6,0) as AccClose12Amnt6
                                            ,ifnull(AccClose24Amnt6,0) as AccClose24Amnt6
                                            ,ifnull(Installment6Amnt1,0) as Installment6Amnt1 
                                            ,ifnull(Installment6Amnt2,0) as Installment6Amnt2 
                                            ,ifnull(Installment6Amnt3,0) as Installment6Amnt3 
                                            ,ifnull(Installment6Amnt4,0) as Installment6Amnt4 
                                            ,ifnull(Installment6Amnt5,0) as Installment6Amnt5 
                                            ,ifnull(Installment6Amnt6,0) as Installment6Amnt6 
                                            ,ifnull(Installment12Amnt1,0) as Installment12Amnt1
                                            ,ifnull(Installment12Amnt2,0) as Installment12Amnt2
                                            ,ifnull(Installment12Amnt3,0) as Installment12Amnt3
                                            ,ifnull(Installment12Amnt4,0) as Installment12Amnt4
                                            ,ifnull(Installment12Amnt5,0) as Installment12Amnt5
                                            ,ifnull(Installment12Amnt6,0) as Installment12Amnt6
                                            ,ifnull(Installment24Amnt1,0) as Installment24Amnt1
                                            ,ifnull(Installment24Amnt2,0) as Installment24Amnt2
                                            ,ifnull(Installment24Amnt3,0) as Installment24Amnt3
                                            ,ifnull(Installment24Amnt4,0) as Installment24Amnt4
                                            ,ifnull(Installment24Amnt5,0) as Installment24Amnt5
                                            ,ifnull(Installment24Amnt6,0) as Installment24Amnt6
                                            FROM CPSPayAmnt
                                            WHERE CustomerID = '{0}' ", customerwhere);
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        FestCustom customdata = new FestCustom();
                        customdata.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        customdata.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                        customdata.CaseID = reader.GetString("CaseID") ?? "";
                        customdata.WorkNo = reader.GetString("WorkNo") ?? "";
                        customdata.LedNumber = reader.GetString("LedNumber") ?? "";
                        customdata.CardNo1 = reader.GetString("CardNo1") ?? "";
                        customdata.AccCloseAmnt1 = reader.GetDouble("AccCloseAmnt1");
                        customdata.AccClose6Amnt1 = reader.GetDouble("AccClose6Amnt1");
                        customdata.AccClose12Amnt1 = reader.GetDouble("AccClose12Amnt1");
                        customdata.AccClose24Amnt1 = reader.GetDouble("AccClose24Amnt1");
                        customdata.CardNo2 = reader.GetString("CardNo2") ?? "";
                        customdata.AccCloseAmnt2 = reader.GetDouble("AccCloseAmnt2");
                        customdata.AccClose6Amnt2 = reader.GetDouble("AccClose6Amnt2");
                        customdata.AccClose12Amnt2 = reader.GetDouble("AccClose12Amnt2");
                        customdata.AccClose24Amnt2 = reader.GetDouble("AccClose24Amnt2");
                        customdata.CardNo3 = reader.GetString("CardNo3") ?? "";
                        customdata.AccCloseAmnt3 = reader.GetDouble("AccCloseAmnt3");
                        customdata.AccClose6Amnt3 = reader.GetDouble("AccClose6Amnt3");
                        customdata.AccClose12Amnt3 = reader.GetDouble("AccClose12Amnt3");
                        customdata.AccClose24Amnt3 = reader.GetDouble("AccClose24Amnt3");
                        customdata.CardNo4 = reader.GetString("CardNo4") ?? "";
                        customdata.AccCloseAmnt4 = reader.GetDouble("AccCloseAmnt4");
                        customdata.AccClose6Amnt4 = reader.GetDouble("AccClose6Amnt4");
                        customdata.AccClose12Amnt4 = reader.GetDouble("AccClose12Amnt4");
                        customdata.AccClose24Amnt4 = reader.GetDouble("AccClose24Amnt4");
                        customdata.CardNo5 = reader.GetString("CardNo5") ?? "";
                        customdata.AccCloseAmnt5 = reader.GetDouble("AccCloseAmnt5");
                        customdata.AccClose6Amnt5 = reader.GetDouble("AccClose6Amnt5");
                        customdata.AccClose12Amnt5 = reader.GetDouble("AccClose12Amnt5");
                        customdata.AccClose24Amnt5 = reader.GetDouble("AccClose24Amnt5");
                        customdata.CardNo6 = reader.GetString("CardNo6") ?? "";
                        customdata.AccCloseAmnt6 = reader.GetDouble("AccCloseAmnt6");
                        customdata.AccClose6Amnt6 = reader.GetDouble("AccClose6Amnt6");
                        customdata.AccClose12Amnt6 = reader.GetDouble("AccClose12Amnt6");
                        customdata.AccClose24Amnt6 = reader.GetDouble("AccClose24Amnt6");
                        customdata.Installment6Amnt1   = reader.GetDouble("Installment6Amnt1");
                        customdata.Installment6Amnt2   = reader.GetDouble("Installment6Amnt2");
                        customdata.Installment6Amnt3   = reader.GetDouble("Installment6Amnt3");
                        customdata.Installment6Amnt4   = reader.GetDouble("Installment6Amnt4");
                        customdata.Installment6Amnt5   = reader.GetDouble("Installment6Amnt5");
                        customdata.Installment6Amnt6   = reader.GetDouble("Installment6Amnt6");
                        customdata.Installment12Amnt1  = reader.GetDouble("Installment12Amnt1");
                        customdata.Installment12Amnt2  = reader.GetDouble("Installment12Amnt2");
                        customdata.Installment12Amnt3  = reader.GetDouble("Installment12Amnt3");
                        customdata.Installment12Amnt4  = reader.GetDouble("Installment12Amnt4");
                        customdata.Installment12Amnt5  = reader.GetDouble("Installment12Amnt5");
                        customdata.Installment12Amnt6  = reader.GetDouble("Installment12Amnt6");
                        customdata.Installment24Amnt1  = reader.GetDouble("Installment24Amnt1");
                        customdata.Installment24Amnt2  = reader.GetDouble("Installment24Amnt2");
                        customdata.Installment24Amnt3  = reader.GetDouble("Installment24Amnt3");
                        customdata.Installment24Amnt4  = reader.GetDouble("Installment24Amnt4");
                        customdata.Installment24Amnt5  = reader.GetDouble("Installment24Amnt5");
                        customdata.Installment24Amnt6 = reader.GetDouble("Installment24Amnt6");
                        customdataall.Add(customdata);
                    }
                }
                connection.Close();
            }
            return customdataall;
        }
        public List<DataCPSPerson> doCheckCPSCustomNotInMaster()
        {
            List<DataCPSPerson> customdataall = new List<DataCPSPerson>();
            if (connection != null)
            {
                connection.Open();
                var sqlcmd = @" SELECT
                                ifnull(WorkNo,'') as WorkNo
                                ,ifnull(LedNumber,'') as LedNumber
                                ,ifnull(CustomerID,'') as CustomerID
                                ,ifnull(CustomerName,'') as CustomerName
                            FROM CPSPayAmnt
                            WHERE CustomerID NOT IN ( SELECT CustomerID FROM DataCPSMaster) 
                            ORDER by LedNumber; ";

                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        DataCPSPerson customdata = new DataCPSPerson();
                        customdata.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        customdata.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                        customdata.WorkNo = reader.GetString("WorkNo") ?? "";
                        customdata.LedNumber = reader.GetString("LedNumber") ?? "";
                        customdataall.Add(customdata);
                    }
                }
                connection.Close();
            }
            return customdataall;
        }
        public List<FestCustom> doGetCustomDuplicateIDAll()
        {
            List<FestCustom> customdataall = new List<FestCustom>();
            if (connection != null)
            {
                connection.Open();
                var sqlcmd = string.Format(@" SELECT
                                               ifnull(WorkNo,'') as WorkNo
                                              ,ifnull(LedNumber,'') as LedNumber
                                              ,ifnull(CustomerID,'') as CustomerID
                                              ,ifnull(CustomerName,'') as CustomerName
                                              ,ifnull(CardNo1,'') as CardNo1
                                              ,ifnull(AccCloseAmnt1,0) as AccCloseAmnt1
                                              ,ifnull(AccClose6Amnt1,0) as AccClose6Amnt1
                                              ,ifnull(AccClose12Amnt1,0) as AccClose12Amnt1
                                              ,ifnull(AccClose24Amnt1,0) as AccClose24Amnt1
                                              ,ifnull(CardNo2,'') as CardNo2
                                              ,ifnull(AccCloseAmnt2,0) as AccCloseAmnt2
                                              ,ifnull(AccClose6Amnt2,0) as AccClose6Amnt2
                                              ,ifnull(AccClose12Amnt2,0) as AccClose12Amnt2
                                              ,ifnull(AccClose24Amnt2,0) as AccClose24Amnt2
                                              ,ifnull(CardNo3,'') as CardNo3
                                              ,ifnull(AccCloseAmnt3,0) as AccCloseAmnt3
                                              ,ifnull(AccClose6Amnt3,0) as AccClose6Amnt3
                                              ,ifnull(AccClose12Amnt3,0) as AccClose12Amnt3
                                              ,ifnull(AccClose24Amnt3,0) as AccClose24Amnt3
                                              ,ifnull(CardNo4,'') as CardNo4
                                              ,ifnull(AccCloseAmnt4,0) as AccCloseAmnt4
                                              ,ifnull(AccClose6Amnt4,0) as AccClose6Amnt4
                                              ,ifnull(AccClose12Amnt4,0) as AccClose12Amnt4
                                              ,ifnull(AccClose24Amnt4,0) as AccClose24Amnt4
                                              ,ifnull(CardNo5,'') as CardNo5
                                              ,ifnull(AccCloseAmnt5,0) as AccCloseAmnt5
                                              ,ifnull(AccClose6Amnt5,0) as AccClose6Amnt5
                                              ,ifnull(AccClose12Amnt5,0) as AccClose12Amnt5
                                              ,ifnull(AccClose24Amnt5,0) as AccClose24Amnt5
                                              ,ifnull(CardNo6,'') as CardNo6
                                              ,ifnull(AccCloseAmnt6,0) as AccCloseAmnt6
                                              ,ifnull(AccClose6Amnt6,0) as AccClose6Amnt6
                                              ,ifnull(AccClose12Amnt6,0) as AccClose12Amnt6
                                              ,ifnull(AccClose24Amnt6,0) as AccClose24Amnt6
                                            FROM CPSPayAmnt
                                            WHERE CustomerID in (SELECT CustomerID
                                                    FROM CPSPayAmnt
					                                GROUP BY CustomerID
					                                HAVING COUNT(CustomerID) > 1)
                                            ORDER by CustomerID;");
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        FestCustom customdata = new FestCustom();
                        customdata.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        customdata.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");

                        customdata.WorkNo = reader.GetString("WorkNo") ?? "";
                        customdata.LedNumber = reader.GetString("LedNumber") ?? "";
                        customdata.CardNo1 = reader.GetString("CardNo1") ?? "";
                        customdata.AccCloseAmnt1 = reader.GetDouble("AccCloseAmnt1");
                        customdata.AccClose6Amnt1 = reader.GetDouble("AccClose6Amnt1");
                        customdata.AccClose12Amnt1 = reader.GetDouble("AccClose12Amnt1");
                        customdata.AccClose24Amnt1 = reader.GetDouble("AccClose24Amnt1");
                        customdata.CardNo2 = reader.GetString("CardNo2") ?? "";
                        customdata.AccCloseAmnt2 = reader.GetDouble("AccCloseAmnt2");
                        customdata.AccClose6Amnt2 = reader.GetDouble("AccClose6Amnt2");
                        customdata.AccClose12Amnt2 = reader.GetDouble("AccClose12Amnt2");
                        customdata.AccClose24Amnt2 = reader.GetDouble("AccClose24Amnt2");
                        customdata.CardNo3 = reader.GetString("CardNo3") ?? "";
                        customdata.AccCloseAmnt3 = reader.GetDouble("AccCloseAmnt3");
                        customdata.AccClose6Amnt3 = reader.GetDouble("AccClose6Amnt3");
                        customdata.AccClose12Amnt3 = reader.GetDouble("AccClose12Amnt3");
                        customdata.AccClose24Amnt3 = reader.GetDouble("AccClose24Amnt3");
                        customdata.CardNo4 = reader.GetString("CardNo4") ?? "";
                        customdata.AccCloseAmnt4 = reader.GetDouble("AccCloseAmnt4");
                        customdata.AccClose6Amnt4 = reader.GetDouble("AccClose6Amnt4");
                        customdata.AccClose12Amnt4 = reader.GetDouble("AccClose12Amnt4");
                        customdata.AccClose24Amnt4 = reader.GetDouble("AccClose24Amnt4");
                        customdata.CardNo5 = reader.GetString("CardNo5") ?? "";
                        customdata.AccCloseAmnt5 = reader.GetDouble("AccCloseAmnt5");
                        customdata.AccClose6Amnt5 = reader.GetDouble("AccClose6Amnt5");
                        customdata.AccClose12Amnt5 = reader.GetDouble("AccClose12Amnt5");
                        customdata.AccClose24Amnt5 = reader.GetDouble("AccClose24Amnt5");
                        customdata.CardNo6 = reader.GetString("CardNo6") ?? "";
                        customdata.AccCloseAmnt6 = reader.GetDouble("AccCloseAmnt6");
                        customdata.AccClose6Amnt6 = reader.GetDouble("AccClose6Amnt6");
                        customdata.AccClose12Amnt6 = reader.GetDouble("AccClose12Amnt6");
                        customdata.AccClose24Amnt6 = reader.GetDouble("AccClose24Amnt6");
                        customdataall.Add(customdata);
                    }
                }
                connection.Close();
            }
            return customdataall;
        }
        public List<DataCPSPerson> doGetFestDuplicateIDAll()
        {
            List<DataCPSPerson> customdataall = new List<DataCPSPerson>();
            if (connection != null)
            {
                connection.Open();
                var sqlcmd = string.Format(@" SELECT
                                               ifnull(WorkNo,'') as WorkNo
                                              ,ifnull(LedNumber,'') as LedNumber
                                              ,ifnull(CustomerID,'') as CustomerID
                                              ,ifnull(CustomerName,'') as CustomerName
                                              ,ifnull(LegalExecRemark,'') as LegalExecRemark
                                            FROM CPSFestData
                                            WHERE CustomerID in (SELECT CustomerID
                                                    FROM CPSFestData
					                                GROUP BY CustomerID
					                                HAVING COUNT(CustomerID) > 1)
                                            ORDER by CustomerID;");
                using var command = new SQLiteCommand(sqlcmd, connection);
                using var reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        DataCPSPerson customdata = new DataCPSPerson();
                        customdata.CustomerID = security_.DecryptString(reader.GetString("CustomerID") ?? "");
                        customdata.CustomerName = security_.DecryptString(reader.GetString("CustomerName") ?? "");
                        customdata.WorkNo = reader.GetString("WorkNo") ?? "";
                        customdata.LedNumber = reader.GetString("LedNumber") ?? "";
                        customdata.LegalExecRemark = reader.GetString("LegalExecRemark") ?? "";
                        customdataall.Add(customdata);
                    }
                }
                connection.Close();
            }
            return customdataall;
        }
        #endregion
        #region Result
        private void doSetResultData(ref DataTable dtresult,string customerid,string result,string remark)
        {
            DataRow drresult = dtresult.NewRow();
            drresult["CustomerID"] = customerid;
            drresult["Result"] = result;
            drresult["Remark"] = remark;
            dtresult.Rows.Add(drresult);
        }
        public DataTable getResultData()
        {
            if(dtResult == null) dtResult = dtService.doCreateResultDataTable();
            return dtResult;
        }
        #endregion
        #region Queuery Data
        public List<DataCPSPerson> getDataWithDataMasterQuery(string sql_cmd)
        {
            List<DataCPSPerson> cardall = new List<DataCPSPerson>();
            if (string.IsNullOrEmpty(sql_cmd)) return cardall;
            sql_cmd = sql_cmd.ToLower();
            string str_where = doGetWhereQuery(sql_cmd);
            if (connection != null)
            {
                connection.Open();
                using (SQLiteCommand command = new SQLiteCommand(str_where, connection))
                {
                    using (SQLiteDataReader reader = command.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                DataCPSPerson carddat = new DataCPSPerson();
                                if (!reader.IsDBNull("CardNo1")) carddat.CardNo1 = reader.GetString("CardNo1");
                                if (!reader.IsDBNull("JudgmentAmnt1")) carddat.JudgmentAmnt1 = reader.GetDouble("JudgmentAmnt1");
                                if (!reader.IsDBNull("PrincipleAmnt1")) carddat.PrincipleAmnt1 = reader.GetDouble("PrincipleAmnt1");
                                if (!reader.IsDBNull("PayAfterJudgAmt1")) carddat.PayAfterJudgAmt1 = reader.GetDouble("PayAfterJudgAmt1");
                                if (!reader.IsDBNull("DeptAmnt1")) carddat.DeptAmnt1 = reader.GetDouble("DeptAmnt1");
                                if (!reader.IsDBNull("LastPayDate1")) carddat.LastPayDate1 = reader.GetString("LastPayDate1");

                                if (!reader.IsDBNull("CardNo2"))carddat.CardNo2 = reader.GetString("CardNo2");
                                if (!reader.IsDBNull("JudgmentAmnt2"))carddat.JudgmentAmnt2 = reader.GetDouble("JudgmentAmnt2");
                                if (!reader.IsDBNull("PrincipleAmnt2"))carddat.PrincipleAmnt2 = reader.GetDouble("PrincipleAmnt2");
                                if (!reader.IsDBNull("PayAfterJudgAmt2"))carddat.PayAfterJudgAmt2 = reader.GetDouble("PayAfterJudgAmt2");
                                if (!reader.IsDBNull("DeptAmnt2"))carddat.DeptAmnt2 = reader.GetDouble("DeptAmnt2");
                                if (!reader.IsDBNull("LastPayDate2"))carddat.LastPayDate2 = reader.GetString("LastPayDate2");

                                if (!reader.IsDBNull("CardNo3"))carddat.CardNo3 = reader.GetString("CardNo3");
                                if (!reader.IsDBNull("JudgmentAmnt3"))carddat.JudgmentAmnt3 = reader.GetDouble("JudgmentAmnt3");
                                if (!reader.IsDBNull("PrincipleAmnt3"))carddat.PrincipleAmnt3 = reader.GetDouble("PrincipleAmnt3");
                                if (!reader.IsDBNull("PayAfterJudgAmt3"))carddat.PayAfterJudgAmt3 = reader.GetDouble("PayAfterJudgAmt3");
                                if (!reader.IsDBNull("DeptAmnt3"))carddat.DeptAmnt3 = reader.GetDouble("DeptAmnt3");
                                if (!reader.IsDBNull("LastPayDate3"))carddat.LastPayDate3 = reader.GetString("LastPayDate3");

                                if (!reader.IsDBNull("CardNo4"))carddat.CardNo4 = reader.GetString("CardNo4");
                                if (!reader.IsDBNull("JudgmentAmnt4"))carddat.JudgmentAmnt4 = reader.GetDouble("JudgmentAmnt4");
                                if (!reader.IsDBNull("PrincipleAmnt4"))carddat.PrincipleAmnt4 = reader.GetDouble("PrincipleAmnt4");
                                if (!reader.IsDBNull("PayAfterJudgAmt4"))carddat.PayAfterJudgAmt4 = reader.GetDouble("PayAfterJudgAmt4");
                                if (!reader.IsDBNull("DeptAmnt4"))carddat.DeptAmnt4 = reader.GetDouble("DeptAmnt4");
                                if (!reader.IsDBNull("LastPayDate4"))carddat.LastPayDate4 = reader.GetString("LastPayDate4");

                                if (!reader.IsDBNull("CardNo5"))carddat.CardNo5 = reader.GetString("CardNo5");
                                if (!reader.IsDBNull("JudgmentAmnt5"))carddat.JudgmentAmnt5 = reader.GetDouble("JudgmentAmnt5");
                                if (!reader.IsDBNull("PrincipleAmnt5"))carddat.PrincipleAmnt5 = reader.GetDouble("PrincipleAmnt5");
                                if (!reader.IsDBNull("PayAfterJudgAmt5"))carddat.PayAfterJudgAmt5 = reader.GetDouble("PayAfterJudgAmt5");
                                if (!reader.IsDBNull("DeptAmnt5"))carddat.DeptAmnt5 = reader.GetDouble("DeptAmnt5");
                                if (!reader.IsDBNull("LastPayDate5"))carddat.LastPayDate5 = reader.GetString("LastPayDate5");

                                if (!reader.IsDBNull("CardNo6"))carddat.CardNo6 = reader.GetString("CardNo6");
                                if (!reader.IsDBNull("JudgmentAmnt6"))carddat.JudgmentAmnt6 = reader.GetDouble("JudgmentAmnt6");
                                if (!reader.IsDBNull("PrincipleAmnt6"))carddat.PrincipleAmnt6 = reader.GetDouble("PrincipleAmnt6");
                                if (!reader.IsDBNull("PayAfterJudgAmt6"))carddat.PayAfterJudgAmt6 = reader.GetDouble("PayAfterJudgAmt6");
                                if (!reader.IsDBNull("DeptAmnt6"))carddat.DeptAmnt6 = reader.GetDouble("DeptAmnt6");
                                if (!reader.IsDBNull("LastPayDate6"))carddat.LastPayDate6 = reader.GetString("LastPayDate6");

                                if (!reader.IsDBNull("CustomerName"))carddat.CustomerName = security_.DecryptString(reader.GetString("CustomerName"));
                                if (!reader.IsDBNull("CustomerID"))carddat.CustomerID = security_.DecryptString(reader.GetString("CustomerID"));
                                if (!reader.IsDBNull("CustomerTel")) carddat.CustomerTel = security_.DecryptString(reader.GetString("CustomerTel"));

                                if (!reader.IsDBNull("LegalStatus"))carddat.LegalStatus = reader.GetString("LegalStatus");
                                if (!reader.IsDBNull("BlackNo"))carddat.BlackNo = reader.GetString("BlackNo");
                                if (!reader.IsDBNull("RedNo"))carddat.RedNo = reader.GetString("RedNo");
                                if (!reader.IsDBNull("JudgeDate"))carddat.JudgeDate = reader.GetString("JudgeDate");
                                if (!reader.IsDBNull("CourtName"))carddat.CourtName = reader.GetString("CourtName");
                                if (!reader.IsDBNull("CollectorName"))carddat.CollectorName = reader.GetString("CollectorName");
                                if (!reader.IsDBNull("CollectorTeam"))carddat.CollectorTeam = reader.GetString("CollectorTeam");
                                if (!reader.IsDBNull("CollectorTel"))carddat.CollectorTel = reader.GetString("CollectorTel");
                                if (!reader.IsDBNull("LegalExecDate"))carddat.LegalExecDate = reader.GetString("LegalExecDate");
                                if (!reader.IsDBNull("LegalExecRemark"))carddat.LegalExecRemark = reader.GetString("LegalExecRemark");
                                if (!reader.IsDBNull("CustomFlag")) carddat.CustomFlag = reader.GetString("CustomFlag");
                                cardall.Add(carddat);
                            }
                        }

                    }
                }
                connection.Close();
            }
            return cardall;
        }

        private string doGetWhereQuery(string sql_cmd)
        {
            string[] whereencrypt = new string[] { "customerid", "customername", "customertel" };
            string sql_return = sql_cmd;
            int indexwhere = sql_cmd.IndexOf("where");
            if(indexwhere > 0)
            {
                for(int i = 0; i < whereencrypt.Length; i++)
                {
                    if (sql_cmd.Contains(whereencrypt[i]))
                    {
                      string paramencrpt = doGetEncpitpDataWhere(sql_cmd.Substring(indexwhere), whereencrypt[0]);
                        if (!string.IsNullOrEmpty(paramencrpt))
                        {
                            string encriptstr = security_.EncryptString(paramencrpt);
                            sql_return = sql_cmd.Replace(paramencrpt, encriptstr);
                        }
                    }
                }
                
                
            }
            return sql_return;
        }
        private string doGetEncpitpDataWhere(string sqlwhere,string keycolname)
        {
            int index_strat_ = -1;
            int index_end_ = -1;
            int index_start = sqlwhere.IndexOf(keycolname);
            for (int i = 0; i < sqlwhere.Length; i++)
            {
                string chardat = string.Format("'");
                if (sqlwhere[i] == chardat[0])
                {
                    if (index_strat_ == -1)
                    {
                        index_strat_ = i;
                        continue;
                    }
                    if (index_end_ == -1 && index_strat_ != -1)
                    { 
                        index_end_ = i; 
                        break;
                    }
                }
            }
            if(index_strat_ >= 0 && index_end_ >= 0)
            {
                return sqlwhere.Substring(index_strat_ + 1, (index_end_ - index_strat_)-1);
            }
            return string.Empty;
        }
        #endregion
    }
   
}
