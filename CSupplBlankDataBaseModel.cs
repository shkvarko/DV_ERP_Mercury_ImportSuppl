using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ERPMercury.WebAPI.Data.Classes;

namespace ERPMercuryImportSuppl
{
    public static class CSupplBlankDataBaseModel
    {
        /// <summary>
        /// Возвращает заголовок заказа
        /// </summary>
        /// <param name="objProfile">профайл</param>
        /// <param name="DeliveryDate">дата доставки</param>
        /// <param name="CustomerId">код клиента в IB</param>
        /// <param name="ChildDepartCode">код дочернего клиента</param>
        /// <param name="DepartCode">код подразделения</param>
        /// <param name="IsBonus">признак "бонусный заказ"</param>
        /// <param name="RttCode">код РТТ</param>
        /// <param name="Description">примечание</param>
        /// <param name="SqlCmd">SQL-команда</param>
        /// <param name="strErr">текст ошибки</param>
        /// <returns>объект класса "Order"</returns>
        public static Order CreateOrderHeader( UniXP.Common.CProfile objProfile, System.DateTime DeliveryDate, System.Int32 CustomerId, System.String ChildDepartCode,
            System.String DepartCode, System.Boolean IsBonus,
            System.String RttCode, System.String Description, System.Data.SqlClient.SqlCommand SqlCmd, ref System.String strErr)
        {
            Order objOrder = null;

            try
            {
                System.Data.SqlClient.SqlCommand cmd = null;
                if (SqlCmd == null)
                {
                    System.Data.SqlClient.SqlConnection DBConnection = objProfile.GetDBSource();
                    if (DBConnection == null)
                    {
                        strErr += ("\nОтсутствует соединение с БД.");
                        return objOrder;
                    }

                    cmd = new System.Data.SqlClient.SqlCommand();
                    cmd.Connection = DBConnection;
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                }
                else
                {
                    cmd = SqlCmd;
                    cmd.Parameters.Clear();
                }


                cmd.CommandText = System.String.Format("[{0}].[dbo].[usp_ConvertSupplFromExcel]", objProfile.GetOptionsDllDBName());
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@CustomerId", System.Data.SqlDbType.Int));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@DepartCode", System.Data.SqlDbType.NVarChar, 3));

                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Suppl_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Depart_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Customer_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@CustomerChild_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Rtt_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Address_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@SupplType_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));

                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ERROR_NUM", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ERROR_MES", System.Data.SqlDbType.NVarChar, 4000));
                cmd.Parameters["@ERROR_MES"].Direction = System.Data.ParameterDirection.Output;


                cmd.Parameters["@CustomerId"].Value = CustomerId;
                cmd.Parameters["@DepartCode"].Value = DepartCode;

                if (ChildDepartCode.Trim().Length > 0)
                {
                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ChildDepartCode", System.Data.SqlDbType.NVarChar, 56));
                    cmd.Parameters["@ChildDepartCode"].Value = ChildDepartCode;
                }
                if (RttCode.Trim().Length > 0)
                {
                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@RttCode", System.Data.SqlDbType.NVarChar, 56));
                    cmd.Parameters["@RttCode"].Value = RttCode;
                }
                cmd.ExecuteNonQuery();
                System.Int32 iRes = (System.Int32)cmd.Parameters["@RETURN_VALUE"].Value;

                if (iRes == 0)
                {
                    if ((cmd.Parameters["@Depart_Guid"].Value != System.DBNull.Value) &&
                        (cmd.Parameters["@Customer_Guid"].Value != System.DBNull.Value))
                    {
                        objOrder = new Order();
                        objOrder.id = System.Convert.ToString(cmd.Parameters["@Suppl_Guid"].Value);
                        objOrder.deliveryDate = DeliveryDate;
                        objOrder.beginDate = System.DateTime.Now;
                        objOrder.typeId = ((ChildDepartCode.Trim().Length == 0) ? 0 : 1);
                        objOrder.clientId = System.Convert.ToString(cmd.Parameters["@Customer_Guid"].Value);
                        objOrder.departId = System.Convert.ToString(cmd.Parameters["@Depart_Guid"].Value);
                        objOrder.rttId = System.Convert.ToString(cmd.Parameters["@Rtt_Guid"].Value);
                        objOrder.addressId = System.Convert.ToString(cmd.Parameters["@Address_Guid"].Value);
                        objOrder.bonus = IsBonus;
                        objOrder.description = Description;
                        objOrder.orderGeoPoint = new GeoCoordinate();
                    }
                }
                else
                {
                    strErr += ("\n" + System.Convert.ToString(cmd.Parameters["@ERROR_MES"].Value));
                }
            }
            catch (System.Exception f)
            {
                strErr += ("\n" + f.Message);
            }

            return objOrder;
        }

        /// <summary>
        /// Возвращает табличную часть к заказу
        /// </summary>
        /// <param name="ColumnQty">столбец с количеством</param>
        /// <param name="objProfile">профайл</param>
        /// <param name="objNodesCollection">исходный список позиций для заказа</param>
        /// <param name="SqlCmd">SQL-команда</param>
        /// <param name="strErr">текст ошибки</param>
        /// <returns>список объектов класса "OrderItem"</returns>
        public static List<OrderItem> CreateOrderTablePart(object ColumnQty, UniXP.Common.CProfile objProfile, 
            DevExpress.XtraTreeList.Nodes.TreeListNodes objNodesCollection,
            System.Data.SqlClient.SqlCommand SqlCmd, ref System.String strErr)
        {
            List<OrderItem> objOrderItemList = null;
            if ((objNodesCollection == null) || (objNodesCollection.Count == 0))
            {
                strErr += ("\nВ заказе отсутствует список товаров.");
                return objOrderItemList;
            }

            try
            {
                System.Data.SqlClient.SqlCommand cmd = null;
                if (SqlCmd == null)
                {
                    System.Data.SqlClient.SqlConnection DBConnection = objProfile.GetDBSource();
                    if (DBConnection == null)
                    {
                        strErr += ("\nОтсутствует соединение с БД.");
                        return objOrderItemList;
                    }

                    cmd = new System.Data.SqlClient.SqlCommand();
                    cmd.Connection = DBConnection;
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                }
                else
                {
                    cmd = SqlCmd;
                    cmd.Parameters.Clear();
                }

                cmd.Parameters.Clear();
                cmd.CommandText = System.String.Format("[{0}].[dbo].[usp_ConvertSupplItmsFromExcel]", objProfile.GetOptionsDllDBName());
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@PartsId", System.Data.SqlDbType.Int));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@PartsQty", System.Data.SqlDbType.Int));

                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@SupplItms_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Parts_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Measure_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@SplItms_OrderQty", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@SplItms_Quatity", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@SplItms_Discount", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ERROR_NUM", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ERROR_MES", System.Data.SqlDbType.NVarChar, 4000));
                cmd.Parameters["@ERROR_MES"].Direction = System.Data.ParameterDirection.Output;

                System.Int32 iRes = 0;
                objOrderItemList = new List<OrderItem>();
                OrderItem objOrderItem = null;
                System.Int32 iCurrentIndex = 0;

                foreach (DevExpress.XtraTreeList.Nodes.TreeListNode objNode in objNodesCollection)
                {
                    iCurrentIndex++;
                    if (objNode.Tag == null) { continue; }
                    cmd.Parameters["@PartsId"].Value = System.Convert.ToInt32(objNode.Tag);
                    cmd.Parameters["@PartsQty"].Value = System.Convert.ToInt32(objNode.GetValue(ColumnQty));

                    cmd.ExecuteNonQuery();
                    iRes = (System.Int32)cmd.Parameters["@RETURN_VALUE"].Value;

                    if (iRes == 0)
                    {
                        objOrderItem = new OrderItem();
                        objOrderItem.id = System.Convert.ToString(cmd.Parameters["@SupplItms_Guid"].Value);
                        objOrderItem.orderQuantity = System.Convert.ToInt32(cmd.Parameters["@SplItms_OrderQty"].Value);
                        objOrderItem.quantity = System.Convert.ToInt32(cmd.Parameters["@SplItms_Quatity"].Value);
                        objOrderItem.productId = System.Convert.ToString(cmd.Parameters["@Parts_Guid"].Value);

                        objOrderItemList.Add(objOrderItem);
                    }
                    else
                    {
                        strErr += (String.Format("{0} {1}", iCurrentIndex, (System.String)cmd.Parameters["@ERROR_MES"].Value));
                    }
                }

            }
            catch (System.Exception f)
            {
                strErr += ("\n" + f.Message);
            }

            return objOrderItemList;
        }

        /// <summary>
        /// Возвращает строку подключения к web-сервису
        /// </summary>
        /// <param name="objProfile">профайл</param>
        /// <param name="SqlCmd">SQL-команда</param>
        /// <param name="strErr">текст ошибки</param>
        /// <returns>строка подключения к web-сервису</returns>
        public static System.String GetWebServiceName(UniXP.Common.CProfile objProfile, System.Data.SqlClient.SqlCommand SqlCmd, ref System.String strErr)
        {
            System.String strWebServiceName = System.String.Empty;

            try
            {
                System.Data.SqlClient.SqlCommand cmd = null;
                if (SqlCmd == null)
                {
                    System.Data.SqlClient.SqlConnection DBConnection = objProfile.GetDBSource();
                    if (DBConnection == null)
                    {
                        strErr += ("\nОтсутствует соединение с БД.");
                        return strWebServiceName;
                    }

                    cmd = new System.Data.SqlClient.SqlCommand();
                    cmd.Connection = DBConnection;
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                }
                else
                {
                    cmd = SqlCmd;
                    cmd.Parameters.Clear();
                }

                cmd.Parameters.Clear();
                cmd.CommandText = System.String.Format("[{0}].[dbo].[usp_GetWebServiceName]", objProfile.GetOptionsDllDBName());
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@SoapUrl", System.Data.SqlDbType.NVarChar, 400));
                cmd.Parameters["@SoapUrl"].Direction = System.Data.ParameterDirection.Output;
                cmd.ExecuteNonQuery();

                strWebServiceName = System.Convert.ToString(cmd.Parameters["@SoapUrl"].Value);
            }
            catch (System.Exception f)
            {
                strErr += ("\n" + f.Message);
            }

            return strWebServiceName;
        }

    }
}
