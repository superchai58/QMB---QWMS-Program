using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace QWMS_SendData
{

    class Program
    {
        private static clsDB DB = new clsDB();      

        [Obsolete]
        static void Main(string[] args)
        {
            string strConnFromAppconfig = ConfigurationSettings.AppSettings["strConn"];
            string strConn = QWMS_SendData.EncryptAndDecrypt.Decrypt(strConnFromAppconfig, "RAK");

            QMSSDK.Db.Connections.CreateCn(strConn);

            SFGSendToQWMS_MainProcess();

            ReturnMaterialSendToQWMS_MainProcess();

            ReturnMaterialGetResultFromQWMS_MainProcess();
        }

        private static void SFGSendToQWMS_MainProcess()
        {
            DataTable dtSFGSendData = new DataTable();
            dtSFGSendData = DB.GetSFGSendData("").Tables[0];
            for(int i = 0; i < dtSFGSendData.Rows.Count; i++)
            {
                if (SendDataToQWMS(dtSFGSendData.Rows[i]["API"].ToString(), dtSFGSendData.Rows[i]["XMLString"].ToString(), "QMB_QWMS_QMSDOC", "SP_GetPalletIDData_SMT"))
                {
                    DB.UpdateWHData("","PASS","");
                }
                else
                {
                    DB.UpdateWHData("", "FAIL", "");
                }
            }
        }

        private static void ReturnMaterialSendToQWMS_MainProcess()
        {
            DataTable dtReturnMaterialSendData = new DataTable();
            dtReturnMaterialSendData = DB.GetReturnMaterialSendData("").Tables[0];
            for (int i = 0; i < dtReturnMaterialSendData.Rows.Count; i++)
            {
                if (!SendDataToQWMS(dtReturnMaterialSendData.Rows[i]["API"].ToString(), dtReturnMaterialSendData.Rows[i]["XMLString"].ToString(), "QMB_QWMS_QMSDOC", "SP_GetRefIDData_SMT"))
                {
                    try
                    {
                        //superchai modify  20230314    (Begin)
                        ConnectDBSMT oCon = new ConnectDBSMT();
                        SqlCommand cmd = new SqlCommand();
                        DataTable dt = new DataTable();
                        string referenceIDList = "";

                        cmd.CommandText = @"SELECT distinct ReferenceID
                                        FROM[10.94.7.15].QSMS.dbo.QSMS_DID_ToWH with(nolock)
                                        WHERE ToWHType = 'Return'
                                        and IsGood = 'Y'
                                        and ReferenceID<> ''
                                        and warehouseid<> ''
                                        and status = '2'
                                        and BatchNo = ''";
                        cmd.CommandTimeout = 180;
                        dt = oCon.Query(cmd);

                        if (dt.Rows.Count > 0)
                        {
                            foreach (DataRow row in dt.Rows)
                            {
                                referenceIDList += row["ReferenceID"].ToString().Trim() + ";";
                            }
                        }

                        DB.SendMail("SendQWMSFail", referenceIDList);
                        //superchai modify  20230314    (End)
                    }
                    catch (Exception ex)
                    {}
                }
            }
        }
        private static void ReturnMaterialGetResultFromQWMS_MainProcess()
        {
            DataTable dtReturnMaterialGetData = new DataTable();
            dtReturnMaterialGetData = DB.GetReturnMaterialGetData("").Tables[0];
            String Result;
            for (int i = 0; i < dtReturnMaterialGetData.Rows.Count; i++)
            {
                Result = UpdateDataFromQWMS(dtReturnMaterialGetData.Rows[i]["URL"].ToString(), dtReturnMaterialGetData.Rows[i]["Body"].ToString());
                if (Result == null || Result.Contains("\"N\""))
                {
                    DB.SendMail("GetQWMSResultAPIFail", "");
                }
                else if (Result.Contains("WERKS ERROR!") || Result.Contains("LGORT ERROR!"))
                {
                    DB.SendMail("GetQWMSResultDataError", "");
                }
                else if (Result.Contains("OMBLN"))
                {
                    DB.UpdateDIDToWHData(dtReturnMaterialGetData.Rows[i]["ReferenceID"].ToString());
                }
            }
        }

        public static bool SendDataToQWMS(string APIAdress, string SendMsg, string AliyunTag, string SPName)
        {
            if (APIAdress == "")
            {
                //APIAdress = "http://10.17.30.86/AliyunMQ/api/MQ/SERVICE";

                /*--superchai   Update API link     20220811--*/
                //APIAdress = "https://qcmcec.quantacn.com/AliyunMQ_Intranet/api/MQ/SERVICE";

                try
                {
                    //-----------------superchai Modify get API from DB     20221118    (Begin)-------------------------------
                    DataTable dt = new DataTable();
                    ConnectDBSMT oCon = new ConnectDBSMT();
                    SqlCommand cmd = new SqlCommand();
                    string strAPILink = "";

                    cmd.CommandText = "Select [Value] From mesPE_ProConfig Where Line = 'F7' AND Station = 'QWMS' AND [Session] = 'qcmcec.quantacn.com' AND [Key] = 'IP.API'";
                    cmd.CommandTimeout = 180;
                    dt = oCon.Query(cmd);

                    if (dt.Rows.Count > 0)
                    {
                        strAPILink = dt.Rows[0]["Value"].ToString().Trim();
                    }

                    APIAdress = strAPILink;
                    //-----------------superchai Modify get API from DB     20221118    (End)-------------------------------
                }
                catch (Exception ex)
                {}
            }
            if (AliyunTag == "")
            {
                AliyunTag = "QMB_QWMS_QMSDOC";
            }
            //转码
            if (SendMsg.Length < 100)
            {
                return true;
            }

            byte[] bytes = Encoding.UTF8.GetBytes(SendMsg);
            SendMsg = Convert.ToBase64String(bytes).Replace("+", "%2B");

            //构造POST数据（消息体 & 默认传1 & 消息类型 & 文档名称（不需要可不传））
            string strParam = $"logistics_interface={HttpUtility.UrlEncode(SendMsg)}&data_digest={HttpUtility.UrlEncode("1")}&msg_type={HttpUtility.UrlEncode(AliyunTag)}&msg_remark={HttpUtility.UrlEncode(SPName)}";
            
            string Result = HttpWebRequestByPost(APIAdress, strParam);

            //检查返回结果
            XMLData XResult = new XMLData();
            XResult = XmlHelper.Instance.Deserialize<XMLData>(Result) as XMLData;
            if (XResult.responseItems[0].success.ToUpper().Equals("TRUE"))
            {
                try
                {
                    //--superchai modify 20230314 (Begin)--                
                    ConnectDBSMT oCon = new ConnectDBSMT();
                    SqlCommand cmd = new SqlCommand();

                    cmd.CommandText = "EXEC QWMS_SendRefDataDetail_QMB 'SendRefIDToQWMS', ''";
                    cmd.CommandTimeout = 180;
                    oCon.ExecuteCommand(cmd);
                    //--superchai modify 20230314 (End)--
                }
                catch (Exception ex)
                {}

                return true;
            }
            else
            {
                return false;
                //throw new Exception("$Api发送失败，返回异常信息，{JsonHelper.Instance.ObjectToJson(FClearance)}");
            }
        }

        public static String UpdateDataFromQWMS(string APIAdress, string SendMsg)
        {
            if (APIAdress == "")
            {
                APIAdress = "https://qcmcec.quantacn.com/QWMSApi/api/QcmcQwms/GetSMTReturn_QMB";
            }

            return HttpWebRequestUseJsonByPost(APIAdress, SendMsg);
        }

        public static string HttpWebRequestByPost(string URL, string SendMsg)
        {
            string FResult = string.Empty;
            try
            {
                byte[] bytSend = Encoding.UTF8.GetBytes(SendMsg);
                HttpWebRequest objHttpWebRequest = (HttpWebRequest)WebRequest.Create(URL);

                objHttpWebRequest.Method = "POST";
                //objHttpWebRequest.Timeout = 5000;
                objHttpWebRequest.ContentType = "application/x-www-form-urlencoded;charset=utf-8";
                objHttpWebRequest.ContentLength = bytSend.Length;

                //发送请求
                System.IO.Stream objHttpWriter = objHttpWebRequest.GetRequestStream();
                objHttpWriter.Write(bytSend, 0, bytSend.Length);
                objHttpWriter.Close();

                //获取响应
                HttpWebResponse objHttpWebResponse = (HttpWebResponse)objHttpWebRequest.GetResponse();
                StreamReader objStreamReader = new StreamReader(objHttpWebResponse.GetResponseStream(), Encoding.UTF8);
                FResult = objStreamReader.ReadToEnd();
                objStreamReader.Close();

                try
                {
                    //Add log to record result of connect to MIS (Connected MIS) (Begin)        superchai modify    20230316              
                    ConnectDBSMT oCon = new ConnectDBSMT();
                    SqlCommand cmd = new SqlCommand();

                    cmd.CommandText = "EXEC QWMS_SendRefDataDetail_QMB 'Connected MIS', 'The result connected to MIS is " + FResult + ", of URL: " + URL + "'";
                    cmd.CommandTimeout = 180;
                    oCon.ExecuteCommand(cmd);
                    //Add log to record result of connect to MIS (Connected MIS) (End)        superchai modify    20230316
                }
                catch (Exception ex)
                {}
            }
            catch (WebException ex)
            {
                //throw ex;
                try
                {
                    //Add log to record result of connect to MIS (Not connected MIS) (Begin)      superchai modify    20230316
                    ConnectDBSMT oCon = new ConnectDBSMT();
                    SqlCommand cmd = new SqlCommand();

                    cmd.CommandText = "EXEC QWMS_SendRefDataDetail_QMB 'Not connected MIS', 'Catch message is " + ex.ToString().Trim() + "'";
                    cmd.CommandTimeout = 180;
                    oCon.ExecuteCommand(cmd);
                    //Add log to record result of connect to MIS (Not connected MIS) (End)      superchai modify    20230316
                }
                catch (Exception)
                {}
            }
            catch (Exception ex)
            {
                //throw ex;
                try
                {
                    //Add log to record result of connect to MIS (Not connected MIS) (Begin)      superchai modify    20230316
                    ConnectDBSMT oCon = new ConnectDBSMT();
                    SqlCommand cmd = new SqlCommand();

                    cmd.CommandText = "EXEC QWMS_SendRefDataDetail_QMB 'Not connected MIS', 'Catch message is " + ex.ToString().Trim() + "'";
                    cmd.CommandTimeout = 180;
                    oCon.ExecuteCommand(cmd);
                    //Add log to record result of connect to MIS (Not connected MIS) (End)      superchai modify    20230316
                }
                catch (Exception)
                {}
            }
            return FResult;
        }

        public static string HttpWebRequestUseJsonByPost(string URL, string SendMsg)
        {
            string FResult = string.Empty;
            try
            {
                HttpWebRequest objHttpWebRequest = (HttpWebRequest)WebRequest.Create(URL);

                objHttpWebRequest.Method = "POST";
                //objHttpWebRequest.Timeout = 5000;
                objHttpWebRequest.ContentType = "application/json; charset=utf-8";
                StreamWriter objStreamWriter = new StreamWriter(objHttpWebRequest.GetRequestStream());
                objStreamWriter.Write(SendMsg);
                objStreamWriter.Flush();

                //获取响应
                HttpWebResponse objHttpWebResponse = (HttpWebResponse)objHttpWebRequest.GetResponse();
                StreamReader objStreamReader = new StreamReader(objHttpWebResponse.GetResponseStream(), Encoding.UTF8);
                FResult = objStreamReader.ReadToEnd();
                objStreamReader.Close();
            }
            catch (WebException ex)
            {
                throw ex;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return FResult;
        }

        private static void writeLog(string sLog)
        {
            Console.WriteLine(DateTime.Now.ToString("[" + "HH:mm:ss.FF") + "] " + sLog);
        }

    }
}
