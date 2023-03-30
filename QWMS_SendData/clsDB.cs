using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QWMS_SendData
{
    class clsDB : QMSSDK.Db.WinForm
    {
        internal DataSet GetSFGSendData(string Factory)
        {
            DataSet ds;

            this.spName = "QWMS_SendPalletData_QMB";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@Factory", SqlDbType.VarChar);
            paras[0].Value = Factory;

            ds = this.Execute(paras);
            return ds;
        }

        internal void UpdateWHData(string PalletID, string Result, string Msg)
        {
            this.spName = "QWMS_UpdatePalletDataStatus_QMB";
            SqlParameter[] paras = new SqlParameter[3];
            paras[0] = new SqlParameter("@PalletID", SqlDbType.VarChar);
            paras[0].Value = PalletID;
            paras[1] = new SqlParameter("@Result", SqlDbType.VarChar);
            paras[1].Value = Result;
            paras[2] = new SqlParameter("@Msg", SqlDbType.VarChar);
            paras[2].Value = Msg;

            this.Execute(paras);
        }

        internal DataSet GetReturnMaterialSendData(string Factory)
        {
            DataSet ds;

            this.spName = "QWMS_SendRefData_QMB";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@Factory", SqlDbType.VarChar);
            paras[0].Value = Factory;

            ds =this.Execute(paras);
            return ds;
        }

        internal void UpdateDIDToWHData(String RefID)
        {
            this.spName = "QWMS_UpdateRefDataStatus_QMB";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@RefID", SqlDbType.VarChar);
            paras[0].Value = RefID;

            this.Execute(paras);
        }

        internal DataSet GetReturnMaterialGetData(string v)
        {
            DataSet ds;

            this.spName = "QWMS_GetToWHData_QMB";

            ds = this.Execute();
            return ds;
        }

        //--superchai modify    20230314    (Begin)--
        internal void SendMail(String Type, string referenceID)
        {
            try
            {
                this.spName = "QWMS_SendMail";
                SqlParameter[] paras = new SqlParameter[2];
                paras[0] = new SqlParameter("@Type", SqlDbType.VarChar);
                paras[0].Value = Type;
                paras[1] = new SqlParameter("@ReferenceID", SqlDbType.VarChar);
                paras[1].Value = referenceID;

                this.Execute(paras);
            }
            catch (Exception ex)
            {}
        }
        //--superchai modify    20230314    (End)--
    }
}
