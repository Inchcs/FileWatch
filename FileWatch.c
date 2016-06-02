using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using MySql.Data.MySqlClient;
using System.Xml;
using System.Threading;

//'****************************************************************************
//' Copyright 2015 by ZDT. All Right Reserved.
//'
//' SPI NG 不打件  MAX 10 監控
//' 
//' Date         Version      Author       Comment
//' -------------------------------------------------------------------------
//' 2015/08/16   ZDT-15-001   殷其磊       Initial Release 
//'****************************************************************************

namespace SPI_NG
{
    public partial class Form1 : Form
    {
        //FileInfo fi;       //暫時沒有使用
        //StringBuilder sb;  //暫時沒有使用
        DirectoryInfo dirInfo;
        string title_equ_id = "";
        
        public Form1()
        {
            InitializeComponent();

            string watch_path1 = "";
            string watch_path2 = "";
            string watch_path3 = "";
            string watch_path4 = "";
            string watch_path5 = "";
            string watch_path6 = "";
            string watch_path7 = "";
            string watch_path8 = "";
            string watch_path9 = "";
            string watch_path10 = "";    

            //Load XML File  找出1-10個PATH 設定        
            try            
            {
                XmlDocument xml = new XmlDocument();
                string xmlfile = "Config.xml";
                if (!File.Exists(xmlfile))
                {
                    throw new Exception("XML File not exist! " + xmlfile);
                }
                xml.Load(xmlfile);

                if (xml.GetElementsByTagName("Equ_id").Count > 0)
                {
                    title_equ_id = xml.DocumentElement["Equ_id"].InnerText.Trim();
                }

                if (xml.GetElementsByTagName("path_1").Count > 0)
                {
                    watch_path1 = xml.DocumentElement["path_1"].InnerText.Trim();
                }

                if (xml.GetElementsByTagName("path_2").Count > 0)
                {
                    watch_path2 = xml.DocumentElement["path_2"].InnerText.Trim();
                }

                if (xml.GetElementsByTagName("path_3").Count > 0)
                {
                    watch_path3 = xml.DocumentElement["path_3"].InnerText.Trim();
                }

                if (xml.GetElementsByTagName("path_4").Count > 0)
                {
                    watch_path4 = xml.DocumentElement["path_4"].InnerText.Trim();
                }

                if (xml.GetElementsByTagName("path_5").Count > 0)
                {
                    watch_path5 = xml.DocumentElement["path_5"].InnerText.Trim();
                }

                if (xml.GetElementsByTagName("path_6").Count > 0)
                {
                    watch_path6 = xml.DocumentElement["path_6"].InnerText.Trim();
                }

                if (xml.GetElementsByTagName("path_7").Count > 0)
                {
                    watch_path7 = xml.DocumentElement["path_7"].InnerText.Trim();
                }

                if (xml.GetElementsByTagName("path_8").Count > 0)
                {
                    watch_path8 = xml.DocumentElement["path_8"].InnerText.Trim();
                }

                if (xml.GetElementsByTagName("path_9").Count > 0)
                {
                    watch_path9 = xml.DocumentElement["path_9"].InnerText.Trim();
                }

                if (xml.GetElementsByTagName("path_10").Count > 0)
                {
                    watch_path10 = xml.DocumentElement["path_10"].InnerText.Trim();
                }
            }            
            catch (Exception ex)            
            {
                string errMess = ex.Message;
                MessageBox.Show(errMess);
                System.Environment.Exit(0);                  
            }

            this.Text = "FileWatch Program: " + title_equ_id;

            //如果PATH有內容  則產生watch物件來進行監控 最多10個
            if (watch_path1 != "")
            {
                FileSystemWatcher _watch1 = new FileSystemWatcher();
                MyFileSystemWatcher(_watch1, watch_path1);
            }

            if (watch_path2 != "")
            {
                FileSystemWatcher _watch2 = new FileSystemWatcher();
                MyFileSystemWatcher(_watch2, watch_path2);
            }

            if (watch_path3 != "")
            {
                FileSystemWatcher _watch3 = new FileSystemWatcher();
                MyFileSystemWatcher(_watch3, watch_path3);
            }

            if (watch_path4 != "")
            {
                FileSystemWatcher _watch4 = new FileSystemWatcher();
                MyFileSystemWatcher(_watch4, watch_path4);
            }

            if (watch_path5 != "")
            {
                FileSystemWatcher _watch5 = new FileSystemWatcher();
                MyFileSystemWatcher(_watch5, watch_path5);
            }

            if (watch_path6 != "")
            {
                FileSystemWatcher _watch6 = new FileSystemWatcher();
                MyFileSystemWatcher(_watch6, watch_path6);
            }

            if (watch_path7 != "")
            {
                FileSystemWatcher _watch7 = new FileSystemWatcher();
                MyFileSystemWatcher(_watch7, watch_path7);
            }

            if (watch_path8 != "")
            {
                FileSystemWatcher _watch8 = new FileSystemWatcher();
                MyFileSystemWatcher(_watch8, watch_path8);
            }

            if (watch_path9 != "")
            {
                FileSystemWatcher _watch9 = new FileSystemWatcher();
                MyFileSystemWatcher(_watch9, watch_path9);
            }

            if (watch_path10 != "")
            {
                FileSystemWatcher _watch10 = new FileSystemWatcher();
                MyFileSystemWatcher(_watch10, watch_path10);
            }
        }

        //設定監控物件的屬性
        private void MyFileSystemWatcher(FileSystemWatcher path_watch, string pathinfo)
        {
            try
            {
                //設定所要監控的資料夾
                path_watch.Path = pathinfo;

                //設定所要監控的變更類型
                path_watch.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.CreationTime;

                //設定所要監控的檔案
                path_watch.Filter = "*.ini";

                //設定是否監控子資料夾
                path_watch.IncludeSubdirectories = false;

                //設定是否啟動元件，此部分必須要設定為 true，不然事件是不會被觸發的
                path_watch.EnableRaisingEvents = true;

                //設定觸發事件
                //path_watch.Created += new FileSystemEventHandler(_watch_Created);
                path_watch.Changed += new FileSystemEventHandler(_watch_Created);
                //path_watch.Renamed += new RenamedEventHandler(_watch_Renamed);
                //path_watch.Deleted += new FileSystemEventHandler(_watch_Deleted);
            }
            catch (Exception ex)
            {
                string errMess2 = "監控產生錯誤! 資料夾名稱:" + pathinfo +" 錯誤訊息:"+ ex.Message;
                MessageBox.Show(errMess2);
                System.Environment.Exit(0); 
            }         
        }

        /// <summary>
        /// 當所監控的資料夾有建立ini或是修改時時觸發
        /// </summary>
        private void _watch_Created(object sender, FileSystemEventArgs e)
        {
            string err_bar = "";
            string err_path = "";
            try
            {
                string barcodeID = "";
                string row = "";
                string col = "";
                string rowdata = "";
                string equ_id = "";
                string side = "";
                string pnlsn = "";

                //延時 讓機台檔案產生完畢再行處理 單位是毫秒
                Thread.Sleep(250);

                //Get the File Content and Parent Folder Name
                string strLine;
                dirInfo = new DirectoryInfo(e.FullPath.ToString());
                equ_id = dirInfo.Parent.Name;  //取前1層的資料夾作為機台名稱
                int iTemp = dirInfo.Name.Length;
                if (iTemp >= 3)
                {
                    pnlsn = dirInfo.Name.Substring(0, 3); //取前三碼
                }
                if (iTemp >= 6)
                {
                    side = dirInfo.Name.Substring(dirInfo.Name.Length - 6, 1);  //取第幾面
                    side = side.ToUpper(); 
                }
                err_path = equ_id + ":" + dirInfo.Name;   //Error Message1
                           
                FileStream aFile = new FileStream(dirInfo.FullName, FileMode.Open); 
                StreamReader sr = new StreamReader(aFile); 
                strLine = sr.ReadLine();
                
                while(strLine != null)            
                {
                    strLine = strLine.Trim();
                    string[] sArray;
                    //逐行處理
                    if (strLine.IndexOf("BarcodeID=") >= 0)
                    {
                        sArray = strLine.Split('=');
                        barcodeID = sArray[1].ToUpper();          
                    }

                    if (strLine.IndexOf("Rows") >= 0)
                    {
                        sArray = strLine.Split('=');
                        row = sArray[1];
                    }

                    if (strLine.IndexOf("Columns") >= 0)
                    {
                        sArray = strLine.Split('=');
                        col = sArray[1];
                    }

                    if (strLine.IndexOf("RowData") >= 0)
                    {
                        rowdata = rowdata + " " + strLine;
                    }

                    strLine = sr.ReadLine();            
                }            
                sr.Close();
                err_bar = barcodeID;   //Error Message2

                //確認barcodeID格式 G25-WS9-G20244-1030
                bool bCheck = false;
                string[] atemp = barcodeID.Split('-');

                if (atemp.Length == 4)
                {
                    if (atemp[0].Length == 3 && atemp[1].Length == 3 && atemp[2].Length == 6 && atemp[3].Length == 4)
                    {
                        bCheck = true; 
                    }
                }

                if (bCheck == true)  //正確格式才處理
                {
                    //處理本次量測結果
                    int total = Convert.ToInt32(row) * Convert.ToInt32(col);
                    int iPass = 0;
                    int iNG = 0;
                    rowdata = rowdata.Trim();
                    string[] tempArray = rowdata.Split(' ');

                    rowdata = "";
                    for (int i = 0; i < tempArray.Length; i++)
                    {
                        if (tempArray[i].Trim().IndexOf("__") >= 0)
                        {
                            rowdata = rowdata + "1";
                            iPass++;
                        }

                        if (tempArray[i].Trim().IndexOf("NG") >= 0)
                        {
                            rowdata = rowdata + "0";
                            iNG++;
                        }
                    }
                    rowdata = rowdata.Trim();

                    //Get MySQL Connection
                    string strSQL = "";
                    int TableCount = 0;
                    MySqlConnection mysql = new MySqlConnection();
                    mysql = getMySqlCon();
                    mysql.Open();

                    //Query the dnc.SPI_INI
                    strSQL = "select * from dnc.SPI_INI where BARCODE_ID ='" + barcodeID + "'";
                    MySqlDataAdapter mda = new MySqlDataAdapter(strSQL, mysql);
                    DataSet ds = new DataSet();
                    TableCount = mda.Fill(ds, "new1");

                    if (TableCount > 0)  //已經有資料
                    {
                        string strSQL2 = "";
                        int TableCount2 = 0;
                        string a_result = Convert.ToString(ds.Tables[0].Rows[0][4]);        //取得A面結果
                        string AtoB_result = Convert.ToString(ds.Tables[0].Rows[0][12]);    //取得AtoB面結果

                        //Query the dnc.SPI_SEQUENCE
                        strSQL2 = "select * from dnc.SPI_SEQUENCE where PNL ='" + pnlsn + "' and Count ='" + total.ToString() + "'";
                        MySqlDataAdapter mda2 = new MySqlDataAdapter(strSQL2, mysql);
                        DataSet ds2 = new DataSet();
                        TableCount2 = mda2.Fill(ds2, "new2");   //取得對應關係資料
                        string str_Relation = "";
                        string fin_rowdata = "";

                        //====================================A面===============================//
                        if (side == "A")   //A面直接更新並修改對應關係
                        {
                            if (TableCount2 > 0)
                            {
                                str_Relation = Convert.ToString(ds2.Tables[0].Rows[0][2]);
                                string[] seqArray = str_Relation.Split(',');
                                int itemp = 0;
                                for (int i = 0; i < seqArray.Length; i++) //用AtoB關聯先產生AtoB結果
                                {
                                    itemp = Convert.ToInt32(seqArray[i]);
                                    fin_rowdata = fin_rowdata + rowdata[itemp - 1];
                                }
                                strSQL = "update dnc.SPI_INI set A_NORMAL=" + iPass + ", A_NG=" + iNG + ", A_EQU_ID='" + equ_id + "', A_DEVICE_LAYOUT='" + rowdata + "', A_CREATEDATE=NOW(), AtoB_LAYOUT='" + fin_rowdata + "', AtoB_CREATEDATE=NOW(), TOTAL_DEVICE_LAYOUT='" + rowdata + "', TOTAL_CREATEDATE=NOW()  where BARCODE_ID='" + barcodeID + "'";
                            }
                            else
                            {
                                strSQL = "update dnc.SPI_INI set A_NORMAL=" + iPass + ", A_NG=" + iNG + ", A_EQU_ID='" + equ_id + "', A_DEVICE_LAYOUT='" + rowdata + "', A_CREATEDATE=NOW(), AtoB_LAYOUT='" + rowdata + "', AtoB_CREATEDATE=NOW(), TOTAL_DEVICE_LAYOUT='" + rowdata + "', TOTAL_CREATEDATE=NOW()  where BARCODE_ID='" + barcodeID + "'";
                            }
                            //Execute Non Query
                            MySqlCommand mysqlcom = new MySqlCommand(strSQL, mysql);
                            mysqlcom.ExecuteNonQuery();
                            mysqlcom.Dispose();
                        }
                        //====================================B面===============================//
                        else if (side == "B")   //B面判斷 
                        {
                            if (TableCount2 > 0)  //判斷關係
                            {
                                if (a_result == "")  //A為空 單純只有B面
                                {
                                    strSQL = "update dnc.SPI_INI set B_NORMAL=" + iPass + ", B_NG=" + iNG + ", B_EQU_ID='" + equ_id + "', B_DEVICE_LAYOUT='" + rowdata + "', B_CREATEDATE=NOW(), AtoB_LAYOUT='" + rowdata + "', AtoB_CREATEDATE=NOW(), TOTAL_DEVICE_LAYOUT='" + rowdata + "', TOTAL_CREATEDATE=NOW()  where BARCODE_ID='" + barcodeID + "'";
                                }
                                else
                                {
                                    //Get the Relationship
                                    str_Relation = Convert.ToString(ds2.Tables[0].Rows[0][2]);
                                    fin_rowdata = "";

                                    //Compare the SPI.ini sequence  
                                    string[] seqArray = str_Relation.Split(',');
                                    for (int i = 0; i < seqArray.Length; i++) //用A面結果去跟這次結果比對
                                    {
                                        int itemp = Convert.ToInt32(seqArray[i]);
                                        if (Convert.ToString(a_result[itemp - 1]) == "0" || Convert.ToString(rowdata[i]) == "0")
                                        {
                                            fin_rowdata = fin_rowdata + "0";
                                        }
                                        else
                                        {
                                            fin_rowdata = fin_rowdata + "1";
                                        }
                                    }
                                    strSQL = "update dnc.SPI_INI set B_NORMAL=" + iPass + ", B_NG=" + iNG + ", B_EQU_ID='" + equ_id + "', B_DEVICE_LAYOUT='" + rowdata + "', B_CREATEDATE=NOW(), AtoB_LAYOUT='" + fin_rowdata + "', AtoB_CREATEDATE=NOW(), TOTAL_DEVICE_LAYOUT='" + fin_rowdata + "', TOTAL_CREATEDATE=NOW()  where BARCODE_ID='" + barcodeID + "'";
                                }

                                //Execute Non Query
                                MySqlCommand mysqlcom = new MySqlCommand(strSQL, mysql);
                                mysqlcom.ExecuteNonQuery();
                                mysqlcom.Dispose();
                            }
                        }
                        //====================================C面===============================//
                        else if (side == "C")   //C面判斷 
                        {
                            if (TableCount2 > 0)    //確認C面需要對應關係
                            {
                                if (AtoB_result == "") //只有C面
                                {
                                    strSQL = "update dnc.SPI_INI set C_NORMAL=" + iPass + ", C_NG=" + iNG + ", C_EQU_ID='" + equ_id + "', C_DEVICE_LAYOUT='" + rowdata + "', C_CREATEDATE=NOW(), TOTAL_DEVICE_LAYOUT='" + rowdata + "', TOTAL_CREATEDATE=NOW()  where BARCODE_ID='" + barcodeID + "'";
                                }
                                else
                                {
                                    //Get the Relationship
                                    str_Relation = Convert.ToString(ds2.Tables[0].Rows[0][3]);
                                    fin_rowdata = "";

                                    if (str_Relation.Length > 0)
                                    {
                                        //Compare the SPI.ini sequence
                                        string[] seqArray = str_Relation.Split(',');
                                        for (int i = 0; i < seqArray.Length; i++) //用AtoB結果去跟這次結果比對
                                        {
                                            int itemp = Convert.ToInt32(seqArray[i]);
                                            if (Convert.ToString(AtoB_result[itemp - 1]) == "0" || Convert.ToString(rowdata[i]) == "0")
                                            {
                                                fin_rowdata = fin_rowdata + "0";
                                            }
                                            else
                                            {
                                                fin_rowdata = fin_rowdata + "1";
                                            }
                                        }
                                        strSQL = "update dnc.SPI_INI set C_NORMAL=" + iPass + ", C_NG=" + iNG + ", C_EQU_ID='" + equ_id + "', C_DEVICE_LAYOUT='" + rowdata + "', C_CREATEDATE=NOW(), BtoC_LAYOUT='" + fin_rowdata + "', BtoC_CREATEDATE=NOW(), TOTAL_DEVICE_LAYOUT='" + fin_rowdata + "', TOTAL_CREATEDATE=NOW()  where BARCODE_ID='" + barcodeID + "'";
                                    }
                                    else
                                    {
                                        strSQL = "insert into SPI_LOG (EQP,MSG,STATE) values('" + equ_id + "', 'No CtoB Relation', '0')";
                                    }
                                }
                                //Execute Non Query
                                MySqlCommand mysqlcom = new MySqlCommand(strSQL, mysql);
                                mysqlcom.ExecuteNonQuery();
                                mysqlcom.Dispose();
                            }
                        }

                        ds2.Clear();
                        ds2.Dispose();
                        mda2.Dispose();
                    }
                    else   //沒資料 判斷哪一面
                    {
                        if (side == "A" || side == "B" || side == "C")
                        {
                            string strSQL2 = "";
                            int TableCount2 = 0;

                            int seq_id = 0;
                            string str_Relation = "";

                            //Query the dnc.SPI_SEQUENCE
                            strSQL2 = "select SPI_SEQUENCE_ID, BtoA from dnc.SPI_SEQUENCE where PNL ='" + pnlsn + "' and Count ='" + total.ToString() + "'";
                            MySqlDataAdapter mda3 = new MySqlDataAdapter(strSQL2, mysql);
                            DataSet ds3 = new DataSet();
                            TableCount2 = mda3.Fill(ds3, "new3");

                            //Check the SPI_INI 有沒有對應關係
                            if (TableCount2 > 0)
                            {
                                seq_id = Convert.ToInt32(ds3.Tables[0].Rows[0][0].ToString());
                                str_Relation = Convert.ToString(ds3.Tables[0].Rows[0][1]);
                            }

                            //判斷哪一面 組合SQL
                            if (side == "C")   //C不用AtoB
                            {
                                strSQL = "insert into dnc.SPI_INI (C_EQU_ID, BARCODE_ID, DEVICE_ROW, DEVICE_COLUMN, C_DEVICE_LAYOUT, C_CREATEDATE, TOTAL_DEVICE_LAYOUT, TOTAL_CREATEDATE, SPI_SEQUENCE_ID, C_NORMAL, C_NG) ";
                                strSQL = strSQL + " values('" + equ_id + "', '" + barcodeID + "', " + row + ", " + col + ", '" + rowdata + "', NOW(), '" + rowdata + "', NOW(), " + seq_id + ", " + iPass + ", " + iNG + ") ";
                            }
                            else if (side == "B")  //B不用計算AtoB
                            {
                                strSQL = "insert into dnc.SPI_INI (B_EQU_ID, BARCODE_ID, DEVICE_ROW, DEVICE_COLUMN, B_DEVICE_LAYOUT, B_CREATEDATE, AtoB_LAYOUT, AtoB_CREATEDATE, TOTAL_DEVICE_LAYOUT, TOTAL_CREATEDATE, SPI_SEQUENCE_ID, B_NORMAL, B_NG) ";
                                strSQL = strSQL + " values('" + equ_id + "', '" + barcodeID + "', " + row + ", " + col + ", '" + rowdata + "', NOW(), '" + rowdata + "', NOW(), '" + rowdata + "', NOW(), " + seq_id + ", " + iPass + ", " + iNG + ") ";
                            }
                            else if (side == "A")  //A需要先計算AtoB
                            {
                                strSQL = "insert into dnc.SPI_INI (A_EQU_ID, BARCODE_ID, DEVICE_ROW, DEVICE_COLUMN, A_DEVICE_LAYOUT, A_CREATEDATE, AtoB_LAYOUT, AtoB_CREATEDATE, TOTAL_DEVICE_LAYOUT, TOTAL_CREATEDATE, SPI_SEQUENCE_ID, A_NORMAL, A_NG) ";
                                if (str_Relation.Length > 0)
                                {
                                    string[] atoB_Array = str_Relation.Split(',');
                                    string fin_rowdata = "";
                                    int itemp = 0;
                                    for (int i = 0; i < atoB_Array.Length; i++) //用AtoB關聯先產生AtoB結果
                                    {
                                        itemp = Convert.ToInt32(atoB_Array[i]);
                                        fin_rowdata = fin_rowdata + rowdata[itemp - 1];
                                    }
                                    strSQL = strSQL + " values('" + equ_id + "', '" + barcodeID + "', " + row + ", " + col + ", '" + rowdata + "', NOW(), '" + fin_rowdata + "', NOW(), '" + rowdata + "', NOW(), " + seq_id + ", " + iPass + ", " + iNG + ") ";
                                }
                                else
                                {
                                    strSQL = strSQL + " values('" + equ_id + "', '" + barcodeID + "', " + row + ", " + col + ", '" + rowdata + "', NOW(), '" + rowdata + "', NOW(), '" + rowdata + "', NOW(), " + seq_id + ", " + iPass + ", " + iNG + ") ";
                                }
                            }

                            //Execute Non Query
                            MySqlCommand mysqlcom3 = new MySqlCommand(strSQL, mysql);
                            mysqlcom3.ExecuteNonQuery();
                            mysqlcom3.Dispose();

                            ds3.Clear();
                            ds3.Dispose();
                            mda3.Dispose();
                        }
                    }

                    ds.Clear();
                    ds.Dispose();
                    mda.Dispose();

                    //Close the Mysql Connection
                    mysql.Close();
                    mysql.Dispose();
                }
            }
            catch (Exception ex) 
            {
                string eMess = ex.Message.ToString();  
                MySqlConnection mysql2 = new MySqlConnection();
                mysql2 = getMySqlCon();
                mysql2.Open();

                //Write the Error Log
                string tempSQL = "";
                if (err_bar.Length == 0 && err_path.Length > 0)
                {
                    tempSQL = "insert into SPI_LOG (EQP,MSG,MSG1,STATE) values('OpenFileError', '" + err_path + "', '', '01')";
                }
                else 
                {
                    tempSQL = "insert into SPI_LOG (EQP,MSG,MSG1,STATE) values('OtherError', '" + err_bar + "', '', '0')";
                }
                
                //Execute Non Query
                MySqlCommand mysqlcom2 = new MySqlCommand(tempSQL, mysql2);
                mysqlcom2.ExecuteNonQuery();
                mysqlcom2.Dispose();
                
                //Close the Mysql Connection
                mysql2.Close();
                mysql2.Dispose(); 
            } 
        }

        //MySQL Connection
        //建立數據庫連接
        public static MySqlConnection getMySqlCon()
        {
            XmlDocument xml = new XmlDocument();
            string xmlfile = "Config.xml";
            string strFac = "";
            String mysqlStr = "";

            if (!File.Exists(xmlfile))
            {
                throw new Exception("XML File not exist! " + xmlfile);
            }
            xml.Load(xmlfile);

            if (xml.GetElementsByTagName("FAC").Count > 0)
            {
                strFac = xml.DocumentElement["FAC"].InnerText.Trim();
            }

            //SG
            if (strFac == "SG")
            {
                mysqlStr = "Database=dnc;Data Source=10.182.45.66;User Id=dnc;Password=dncdata;pooling=false;port=3306";
            }
            
            //QHD
            if (strFac == "QHD")
            {
                mysqlStr = "Database=dnc;Data Source=10.86.16.109;User Id=dnc;Password=qhddnc*168;pooling=false;port=3306";
            }
            
            //HA
            if (strFac == "HA")
            {
                mysqlStr = "Database=dnc;Data Source=10.89.164.58;User Id=dnc;Password=dncdata*168;pooling=false;port=3306";
            }
                   
            MySqlConnection mysql = new MySqlConnection(mysqlStr);
            return mysql;
        }

        //關閉程序
        public void button1_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);            
        }

        //暫時沒有使用下列函數
        /*
        /// <summary>
        /// 當所監控的資料夾有文字檔檔案內容有異動時觸發  暫時沒有用
        /// </summary>
        private void _watch_Changed(object sender, FileSystemEventArgs e)
        {
            sb = new StringBuilder();

            dirInfo = new DirectoryInfo(e.FullPath.ToString());

            sb.AppendLine("被異動的檔名為：" + e.Name);
            sb.AppendLine("檔案所在位址為：" + e.FullPath.Replace(e.Name, ""));
            sb.AppendLine("異動內容時間為：" + dirInfo.LastWriteTime.ToString());

            MessageBox.Show(sb.ToString());
        }

        /// <summary>
        /// 當所監控的資料夾有文字檔檔案重新命名時觸發  暫時沒有用
        /// </summary>
        private void _watch_Renamed(object sender, RenamedEventArgs e)
        {
            sb = new StringBuilder();

            fi = new FileInfo(e.FullPath.ToString());

            sb.AppendLine("檔名更新前：" + e.OldName.ToString());
            sb.AppendLine("檔名更新後：" + e.Name.ToString());
            sb.AppendLine("檔名更新前路徑：" + e.OldFullPath.ToString());
            sb.AppendLine("檔名更新後路徑：" + e.FullPath.ToString());
            sb.AppendLine("建立時間：" + fi.LastAccessTime.ToString());

            MessageBox.Show(sb.ToString());
        }

        /// <summary>
        /// 當所監控的資料夾有文字檔檔案有被刪除時觸發  暫時沒有用
        /// </summary>
        private void _watch_Deleted(object sender, FileSystemEventArgs e)
        {
            sb = new StringBuilder();

            sb.AppendLine("被刪除的檔名為：" + e.Name);
            sb.AppendLine("檔案所在位址為：" + e.FullPath.Replace(e.Name, ""));
            sb.AppendLine("刪除時間：" + DateTime.Now.ToString());

            MessageBox.Show(sb.ToString());
        }
        */
    }
}