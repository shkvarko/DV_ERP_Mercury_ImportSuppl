using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ERPMercuryImportSuppl
{
    public class CBlankItem
    {
        #region Переменные, Свойства, Константы
        /// <summary>
        /// Уникальный идентификатор
        /// </summary>
        private System.Int32 m_iPartsId;
        /// <summary>
        /// Уникальный идентификатор
        /// </summary>
        public System.Int32 PartsId
        {
            get { return m_iPartsId; }
            set { m_iPartsId = value; }
        }
        /// <summary>
        /// Наименование товара
        /// </summary>
        private System.String m_strPartsName;
        /// <summary>
        /// Наименование товара
        /// </summary>
        public System.String PartsName
        {
            get { return m_strPartsName; }
            set { m_strPartsName = value; }
        }
        /// <summary>
        /// Артикул товара
        /// </summary>
        private System.String m_strPartsArticle;
        /// <summary>
        /// Артикул товара
        /// </summary>
        public System.String PartsArticle
        {
            get { return m_strPartsArticle; }
            set { m_strPartsArticle = value; }
        }
        /// <summary>
        /// Товарная подгруппа
        /// </summary>
        private System.String m_strPartSubTypeName;
        /// <summary>
        /// Товарная подгруппа
        /// </summary>
        public System.String PartSubTypeName
        {
            get { return m_strPartSubTypeName; }
            set { m_strPartSubTypeName = value; }
        }
        /// <summary>
        /// Товарная группа
        /// </summary>
        private System.String m_strPartTypeName;
        /// <summary>
        /// Товарная группа
        /// </summary>
        public System.String PartTypeName
        {
            get { return m_strPartTypeName; }
            set { m_strPartTypeName = value; }
        }
        /// <summary>
        /// Товарная марка
        /// </summary>
        private System.String m_strProductOwnerName;
        /// <summary>
        /// Товарная марка
        /// </summary>
        public System.String ProductOwnerName
        {
            get { return m_strProductOwnerName; }
            set { m_strProductOwnerName = value; }
        }
        /// <summary>
        /// Цена
        /// </summary>
        private System.Double m_dPrice;
        /// <summary>
        /// Цена
        /// </summary>
        public System.Double Price
        {
            get { return m_dPrice; }
            set { m_dPrice = value; }
        }
        /// <summary>
        /// Остаток
        /// </summary>
        public System.Int32 CurrentQty { get; set; }
        /// <summary>
        /// Резерв
        /// </summary>
        public System.Int32 ReserveQty { get; set; }
        /// <summary>
        /// Остаток + Резерв
        /// </summary>
        public System.Int32 AllQty { get; set; }
        #endregion

        #region Конструктор
        public CBlankItem()
        {
            this.m_iPartsId = 0;
            this.m_strPartsName = "";
            this.m_strPartsArticle = "";
            this.m_strPartSubTypeName = "";
            this.m_strPartTypeName = "";
            this.m_strProductOwnerName = "";
            this.m_dPrice = 0;
            CurrentQty = 0;
            ReserveQty = 0;
            AllQty = 0;
        }
        public CBlankItem(System.Int32 iPartsId, System.String strPartsName, System.String strPartsArticle,
             System.String strPartSubTypeName, System.String strPartTypeName, System.String strProductOwnerName,
             System.Double dPrice)
        {
            this.m_iPartsId = iPartsId;
            this.m_strPartsName = strPartsName;
            this.m_strPartsArticle = strPartsArticle;
            this.m_strPartSubTypeName = strPartSubTypeName;
            this.m_strPartTypeName = strPartTypeName;
            this.m_strProductOwnerName = strProductOwnerName;
            this.m_dPrice = dPrice;
            CurrentQty = 0;
            ReserveQty = 0;
            AllQty = 0;
        }

        #endregion

        #region Список строк для бланка
        /// <summary>
        /// Список товаров для бланка заказа
        /// </summary>
        /// <param name="objProfile">профайл</param>
        /// <param name="DepartTeamGuid">УИ команды</param>
        /// <param name="bWithInfoAboutStock">признак "Запросить остатки на складах отгрузки"</param>
        /// <returns>список позиций для заказа</returns>
        public static List<CBlankItem> GetBlankItemList(UniXP.Common.CProfile objProfile, System.Guid DepartTeamGuid, System.Boolean bWithInfoAboutStock = false)
        {
            List<CBlankItem> objList = new List<CBlankItem>();

            System.Data.SqlClient.SqlConnection DBConnection = objProfile.GetDBSource();
            if (DBConnection == null)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("Не удалось получить список товаров.\nОтсутствует соединение с БД.", "Внимание",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return objList;
            }

            try
            {
                // соединение с БД получено, прописываем команду на выборку данных
                System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
                cmd.Connection = DBConnection;
                cmd.CommandTimeout = 600;
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandText = System.String.Format("[{0}].[dbo].[usp_GetPartsListForBlank]", objProfile.GetOptionsDllDBName());
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ERROR_NUM", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ERROR_MES", System.Data.SqlDbType.NVarChar, 4000));
                cmd.Parameters["@ERROR_MES"].Direction = System.Data.ParameterDirection.Output;
                if (DepartTeamGuid.CompareTo( System.Guid.Empty ) != 0)
                {
                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@DepartTeam_Guid", System.Data.SqlDbType.UniqueIdentifier));
                    cmd.Parameters["@DepartTeam_Guid"].Value = DepartTeamGuid;
                }
                if (bWithInfoAboutStock == true)
                {
                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@GetStock", System.Data.SqlDbType.Bit));
                    cmd.Parameters["@GetStock"].Value = bWithInfoAboutStock;
                }
                System.Data.SqlClient.SqlDataReader rs = cmd.ExecuteReader();
                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        objList.Add(
                            new CBlankItem(
                                (System.Int32)rs["PARTS_ID"], (System.String)rs["PARTS_NAME"], (System.String)rs["PARTS_ARTICLE"],
                                (System.String)rs["PARTSUBTYPE_NAME"], (System.String)rs["PARTTYPE_NAME"], (System.String)rs["OWNER_NAME"],
                                ((rs["Price2"] == System.DBNull.Value) ? 0 : System.Convert.ToDouble(rs["Price2"]))
                                )
                                    {
                                        CurrentQty = System.Convert.ToInt32(rs["CURQTY"]),
                                        ReserveQty = System.Convert.ToInt32(rs["RESQTY"]),
                                        AllQty = System.Convert.ToInt32(rs["QUANTITY"])
                                    }
                                    );
                    }
                }

                rs.Close();
                rs.Dispose();
                cmd.Dispose();
            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("Не удалось получить список товаров.\n\nТекст ошибки: " + f.Message, "Внимание",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
			finally // очищаем занимаемые ресурсы
            {
                DBConnection.Close();
            }
            return objList;
        }
        #endregion

    }
}
