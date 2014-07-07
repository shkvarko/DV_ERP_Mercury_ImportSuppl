using System;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Threading;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
//using ERP_Mercury.Common;

namespace ERPMercuryImportSuppl
{
    public enum enumFormOpenMode
    {
        Unkown = -1,
        CreateBlank = 0,
        ImportSupplFromBlank = 1,
        ImportSupplFromByBarcodes = 2
    }

    public partial class frmSupplBlank : DevExpress.XtraEditors.XtraForm
    {
        #region Свойства, переменные
        private UniXP.Common.CProfile m_objProfile;
        private UniXP.Common.MENUITEM m_objMenuItem;
        private enumFormOpenMode m_enFormOpenMode;
        public delegate void LoadDepartListDelegate(List<ERP_Mercury.Common.CDepart> objDepartList, System.Int32 iRowCountInLis);
        public LoadDepartListDelegate m_LoadDepartListDelegate;
        public System.Threading.Thread ThreadLoadDepartList { get; set; }

        public delegate void LoadDepartTeamListDelegate(List<ERP_Mercury.Common.CDepartTeam> objDepartteamList, System.Int32 iRowCountInLis);
        public LoadDepartTeamListDelegate m_LoadDepartTeamListDelegate;
        public System.Threading.Thread ThreadLoadDepartteamList { get; set; }

        private const System.Int32 iRowsPartForLoadInComboBox = 100;
        private const System.Int32 iThreadSleepTime = 1000;
        private const System.String strWaitLoadList = "ждите... идет заполнение списка";
        private const System.Int32 iMinControlItemHeight = 20;
        private const System.Int32 iSupplState = 5; // заказ создан

        private const System.Int32 iStartRowForImport = 1; 
        private const System.Int32 iColumnBarcode = 1;
        private const System.Int32 iColumnQuantity = 2;
        #endregion

        #region Конструктор
        public frmSupplBlank(UniXP.Common.CProfile objProfile, UniXP.Common.MENUITEM objMenuItem, enumFormOpenMode enFormOpenMode)
        {
            m_objProfile = objProfile;
            m_objMenuItem = objMenuItem;
            m_enFormOpenMode = enFormOpenMode;

            InitializeComponent();
            
            TabControl.ShowTabHeader = DevExpress.Utils.DefaultBoolean.False;

            dtImportSupplDeliveryDate.DateTime = System.DateTime.Today;
            listEditLog.Items.Clear();
            btnImportSupplSave.Enabled = false;
        }
        #endregion

        #region Открытие формы
        private void frmSupplBlank_Load(object sender, EventArgs e)
        {
            try
            {
                switch (m_enFormOpenMode)
                {
                    case enumFormOpenMode.CreateBlank:
                        btnCreateBlank.Enabled = false;
                        btnSaveToDB.Enabled = false;
                        TabControl.SelectedTabPage = tabPageBlank;
                        StartThreadLoadDepartTeamList();
                        break;
                    case enumFormOpenMode.ImportSupplFromBlank:
                        TabControl.SelectedTabPage = tabPageSuppl;
                        break;
                    case enumFormOpenMode.ImportSupplFromByBarcodes:
                        TabControl.SelectedTabPage = tabPageImportSupplByBarCodes;
                        StartThreadLoadDepartList();
                        break;
                    default:
                        break;
                }
            }
            catch (System.Exception f)
            {
                SendMessageToLog("frmSupplBlank_Load. Текст ошибки: " + f.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
            return;

        }
        #endregion

        #region Журнал сообщений
        private void SendMessageToLog(System.String strMessage)
        {
            try
            {
                m_objMenuItem.SimulateNewMessage(strMessage);
            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(
                    "SendMessageToLog.\n\nТекст ошибки: " + f.Message, "Ошибка",
                   System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            return;
        }
        #endregion

        #region Выпадающие списки
        public void StartThreadLoadDepartTeamList()
        {
            try
            {
                // инициализируем делегаты
                m_LoadDepartTeamListDelegate = new LoadDepartTeamListDelegate(LoadDepartTeamListInComboBox);

                cboxDepartTeam.Text = strWaitLoadList;
                cboxDepart.Text = strWaitLoadList;

                cboxDepartTeam.Properties.Items.Clear();
                cboxDepart.Properties.Items.Clear();

                // запуск потока
                this.ThreadLoadDepartteamList = new System.Threading.Thread(LoadDepartTeamListInThread);
                this.ThreadLoadDepartteamList.Start();
            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("StartThreadLoadDepartTeamList().\n\nТекст ошибки: " + f.Message, "Ошибка",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            return;
        }
        /// <summary>
        /// Загружает список команд (метод, выполняемый в потоке)
        /// </summary>
        public void LoadDepartTeamListInThread()
        {
            try
            {
                List<ERP_Mercury.Common.CDepart> objDepartList = ERP_Mercury.Common.CDepart.GetDepartList(m_objProfile);
                List<ERP_Mercury.Common.CDepartTeam> objDepartTeamList = ERP_Mercury.Common.CDepartTeam.GetDepartTeamList(m_objProfile, null, true);

                List<ERP_Mercury.Common.CDepartTeam> objAddDepartTeamList = new List<ERP_Mercury.Common.CDepartTeam>();
                if ((objDepartTeamList != null) && (objDepartTeamList.Count > 0))
                {
                    System.Int32 iRecCount = 0;
                    System.Int32 iRecAllCount = 0;
                    foreach (ERP_Mercury.Common.CDepartTeam objDepartTeam in objDepartTeamList)
                    {
                        if ((objDepartList != null) && (objDepartList.Count > 0))
                        {
                            objDepartTeam.DepartList = objDepartList.Where<ERP_Mercury.Common.CDepart>(x => x.DepartTeam.uuidID.CompareTo(objDepartTeam.uuidID) == 0).ToList<ERP_Mercury.Common.CDepart>();
                        }
                        objAddDepartTeamList.Add(objDepartTeam);
                        iRecCount++;
                        iRecAllCount++;

                        if (iRecCount == iRowsPartForLoadInComboBox)
                        {
                            iRecCount = 0;
                            Thread.Sleep(iThreadSleepTime);
                            this.Invoke(m_LoadDepartTeamListDelegate, new Object[] { objAddDepartTeamList, iRecAllCount });
                            objAddDepartTeamList.Clear();
                        }

                    }
                    if (iRecCount != iRowsPartForLoadInComboBox)
                    {
                        iRecCount = 0;
                        this.Invoke(m_LoadDepartTeamListDelegate, new Object[] { objAddDepartTeamList, iRecAllCount });
                        objAddDepartTeamList.Clear();
                    }

                }

                objDepartTeamList = null;
                objAddDepartTeamList = null;
                this.Invoke(m_LoadDepartTeamListDelegate, new Object[] { null, 0 });
            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("LoadDepartTeamListInThread.\n\nТекст ошибки: " + f.Message, "Ошибка",
                   System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
            }
            return;
        }
        /// <summary>
        /// Загрузка в выпадающий список с командами порции значений
        /// </summary>
        /// <param name="objDepartTeamList">порция из списка команд</param>
        /// <param name="iRowCountInList">всего записей в списке подразделений</param>
        private void LoadDepartTeamListInComboBox(List<ERP_Mercury.Common.CDepartTeam> objDepartTeamList, System.Int32 iRowCountInList)
        {
            try
            {
                if ((objDepartTeamList != null) && (objDepartTeamList.Count > 0) && (cboxImportSupplDepart.Properties.Items.Count < iRowCountInList))
                {
                    cboxDepartTeam.Properties.Items.AddRange(objDepartTeamList);
                }
                else
                {
                    // процесс загрузки данных завершён
                    Thread.Sleep(iThreadSleepTime);

                    cboxDepartTeam.Text = "";
                    cboxDepart.Text = "";

                    cboxDepartTeam.SelectedItem = null;
                    cboxDepart.SelectedItem = null;

                    CheckValidAllParams();
                }

            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("LoadDepartTeamListInComboBox.\n\nТекст ошибки: " + f.Message, "Ошибка",
                   System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
            }
            return;
        }

        /// <summary>
        /// Загружает список подразделений для команды
        /// </summary>
        /// <param name="objDepartTeam"></param>
        private void LoadDepartListForDepartTeam(ERP_Mercury.Common.CDepartTeam objDepartTeam)
        {
            try
            {
                cboxDepart.SelectedItem = null;
                cboxDepart.Properties.Items.Clear();

                if( (objDepartTeam != null) && (objDepartTeam.DepartList != null))
                {
                    cboxDepart.Properties.Items.AddRange(objDepartTeam.DepartList);
                }
            }
            catch (System.Exception f)
            {
                SendMessageToLog("Ошибка загрузки списка подразделений для команды. Текст ошибки: " + f.Message);
            }
            finally
            {
            }
            return;
        }
        /// <summary>
        /// загружает в listbox список назначенных команде товарных марок
        /// </summary>
        /// <param name="objDepartTeam">команда</param>
        private void LoadAssignProductOwner(ERP_Mercury.Common.CDepartTeam objDepartTeam)
        {
            try
            {
                this.tableLayoutPanelDepart.SuspendLayout();
                ((System.ComponentModel.ISupportInitialize)(this.checklstboxProductOwner)).BeginInit();

                checklstboxProductOwner.Items.Clear();

                if ((objDepartTeam != null) && (objDepartTeam.ProductOwnerList != null))
                {
                    foreach (ERP_Mercury.Common.CProductOwner objProductOwner in objDepartTeam.ProductOwnerList)
                    {
                        this.checklstboxProductOwner.Items.Add( objProductOwner, CheckState.Checked );
                    }
                }

                this.tableLayoutPanelDepart.ResumeLayout(false);
                ((System.ComponentModel.ISupportInitialize)(this.checklstboxProductOwner)).EndInit();
            }
            catch (System.Exception f)
            {
                SendMessageToLog("Ошибка загрузки списка назначенных команде товарных марок. Текст ошибки: " + f.Message);
            }
            finally
            {
            }
            return;
        }

        private void cboxDepartTeam_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                LoadDepartListForDepartTeam((ERP_Mercury.Common.CDepartTeam)cboxDepartTeam.SelectedItem);
                LoadAssignProductOwner((ERP_Mercury.Common.CDepartTeam)cboxDepartTeam.SelectedItem);
                LoadCustomerForDepart( null );
                CheckValidAllParams();
            }
            catch (System.Exception f)
            {
                SendMessageToLog("Ошибка выбора команды. Текст ошибки: " + f.Message);
            }
            finally
            {
            }
            return;
        }
        /// <summary>
        /// Загружает список клиентов для подразделения
        /// </summary>
        /// <param name="objDepart">подразделение</param>
        private void LoadCustomerForDepart(ERP_Mercury.Common.CDepart objDepart)
        {
            try
            {
                this.tableLayoutPanelDepart.SuspendLayout();
                ((System.ComponentModel.ISupportInitialize)(this.cboxCustomer.Properties)).BeginInit();
                ((System.ComponentModel.ISupportInitialize)(this.cboxChildCust.Properties)).BeginInit();

                cboxCustomer.SelectedItem = null;
                cboxChildCust.SelectedItem = null;

                cboxCustomer.Properties.Items.Clear();
                cboxChildCust.Properties.Items.Clear();

                if (objDepart != null)
                {
                    if ((objDepart.CustomerList == null) || (objDepart.CustomerList.Count == 0))
                    {
                        objDepart.CustomerList = ERP_Mercury.Common.CCustomer.GetCustomerListWithoutAdvancedProperties(m_objProfile, null, objDepart);
                    }

                    if ((objDepart != null) && (objDepart.CustomerList != null))
                    {
                        cboxCustomer.Properties.Items.AddRange(objDepart.CustomerList);
                    }
                }
                this.tableLayoutPanelDepart.ResumeLayout(false);
                ((System.ComponentModel.ISupportInitialize)(this.cboxCustomer.Properties)).EndInit();
                ((System.ComponentModel.ISupportInitialize)(this.cboxChildCust.Properties)).EndInit();
            }
            catch (System.Exception f)
            {
                SendMessageToLog("Ошибка загрузки списка клиентов для подразделения. Текст ошибки: " + f.Message);
            }
            finally
            {
            }
            return;
        }

        private void cboxDepart_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                ERP_Mercury.Common.CDepart objDepart = (cboxDepart.SelectedItem == null) ? null : (ERP_Mercury.Common.CDepart)cboxDepart.SelectedItem;
                LoadCustomerForDepart(objDepart);
                CheckValidAllParams();
            }
            catch (System.Exception f)
            {
                SendMessageToLog("Ошибка выбора подразделения. Текст ошибки: " + f.Message);
            }
            finally
            {
            }
            return;
        }
        /// <summary>
        /// Загружает список дочерних клиентов для клиента
        /// </summary>
        /// <param name="objCustomer">клиент</param>
        private void LoadChildDepartForCustomer(ERP_Mercury.Common.CCustomer objCustomer)
        {
            try
            {
                this.tableLayoutPanelDepart.SuspendLayout();
                ((System.ComponentModel.ISupportInitialize)(this.cboxChildCust.Properties)).BeginInit();

                cboxChildCust.SelectedItem = null;
                cboxChildCust.Properties.Items.Clear();

                if (objCustomer != null)
                {
                    if ((objCustomer.ChildDepartList == null) || (objCustomer.ChildDepartList.Count == 0))
                    {
                        objCustomer.ChildDepartList = ERP_Mercury.Common.CChildDepart.GetChildDepartList(m_objProfile, null, objCustomer.ID);
                    }

                    if ((objCustomer != null) && (objCustomer.ChildDepartList != null))
                    {
                        foreach (ERP_Mercury.Common.CChildDepart objChildCustomer in objCustomer.ChildDepartList)
                        {
                            cboxChildCust.Properties.Items.Add(objChildCustomer);
                        }
                    }
                }
                this.tableLayoutPanelDepart.ResumeLayout(false);
                ((System.ComponentModel.ISupportInitialize)(this.cboxChildCust.Properties)).EndInit();
            }
            catch (System.Exception f)
            {
                SendMessageToLog("Ошибка загрузки списка дочерних клиентов для клиента. Текст ошибки: " + f.Message);
            }
            finally
            {
            }
            return;
        }

        private void cboxCustomer_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                ERP_Mercury.Common.CCustomer objCustomer = (cboxCustomer.SelectedItem == null) ? null : (ERP_Mercury.Common.CCustomer)cboxCustomer.SelectedItem;
                LoadChildDepartForCustomer(objCustomer );
                CheckValidAllParams();
            }
            catch (System.Exception f)
            {
                SendMessageToLog("Ошибка выбора клиента. Текст ошибки: " + f.Message);
            }
            finally
            {
            }
            return;
        }

        private void CheckValidAllParams()
        {
            try
            {
                btnCreateBlank.Enabled = ((cboxDepart.SelectedItem != null) && (cboxCustomer.SelectedItem != null) && (checklstboxProductOwner.SelectedItems.Count > 0));
                cboxDepartTeam.Properties.Appearance.BackColor = ((cboxDepartTeam.SelectedItem == null) ? System.Drawing.Color.Tomato : System.Drawing.Color.White);
                cboxDepart.Properties.Appearance.BackColor = ((cboxDepart.SelectedItem == null) ? System.Drawing.Color.Tomato : System.Drawing.Color.White);
                cboxCustomer.Properties.Appearance.BackColor = ((cboxCustomer.SelectedItem == null) ? System.Drawing.Color.Tomato : System.Drawing.Color.White);
                //checklstboxProductOwner.Appearance.BackColor = ((checklstboxProductOwner.SelectedItems.Count == 0) ? System.Drawing.Color.Tomato : System.Drawing.Color.White);
            }
            catch (System.Exception f)
            {
                SendMessageToLog("Проверка значений. Текст ошибки: " + f.Message);
            }
            finally
            {
            }
            return;

        }
        #endregion

        #region Формирование бланка

        private void ExportToExcelBlank(ERP_Mercury.Common.CCustomer objCustomer, ERP_Mercury.Common.CChildDepart objChildCustomer,
            ERP_Mercury.Common.CDepart objDepart, List<CBlankItem> objBlankItemList, List<ERP_Mercury.Common.CRtt> objCustomerRttList, 
            System.Boolean bShowStockQty )
        {
            try
            {
                System.String strFileName = (Path.GetTempPath() + "\\" + System.Guid.NewGuid().ToString() + ".xlsx");

                FileInfo newFile = new FileInfo(strFileName);
                if (newFile.Exists)
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(strFileName);
                }
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    // add a new worksheet to the empty workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Suppl");
                    //Add the headers
                    worksheet.Cells[1, 1].Value = "Код клиента";
                    worksheet.Cells[1, 2].Value = "Клиент";
                    worksheet.Cells[1, 3].Value = "Подразделение";
                    worksheet.Cells[1, 4].Value = "Дочерний клиент";
                    worksheet.Cells[1, 5].Value = "Бонус";
                    worksheet.Cells[1, 6].Value = "Примечание";

                    using (var range = worksheet.Cells[1, 1, 1, 6])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Font.Size = 12;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.White);
                        range.Style.Font.Color.SetColor(Color.Black);
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                    }
                    worksheet.Row(1).Height = 20;

                    using (var range = worksheet.Cells[4, 1, 4, 6])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Font.Size = 12;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.White);
                        range.Style.Font.Color.SetColor(Color.Black);
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                    }
                    worksheet.Row(4).Height = 20;

                    worksheet.Cells[2, 6, 2, 6].Style.Locked = false;

                    if (objCustomer != null)
                    {
                        worksheet.Cells[2, 1].Value = objCustomer.InterBaseID;
                        worksheet.Cells[2, 2].Value = objCustomer.FullName;
                    }
                    if (objDepart != null)
                    {
                        worksheet.Cells[2, 3].Value = objDepart.DepartCode;
                    }
                    if (objChildCustomer != null)
                    {
                        worksheet.Cells[2, 4].Value = objChildCustomer.Code;
                    }

                    // попробуем воткнуть розничные точки
                    if ((objCustomerRttList != null) && (objCustomerRttList.Count > 0))
                    {
                        worksheet.Cells[4, 1].Value = "Код РТТ";
                        worksheet.Cells[4, 2].Value = "Наименование РТТ";
                        worksheet.Cells[4, 3].Value = "Адрес";
                        worksheet.Cells[4, 4].Value = "Вкл";

                        for (System.Int32 i = 0; i < objCustomerRttList.Count; i++)
                        {
                            worksheet.Cells[(5 + i), 1].Value = objCustomerRttList[i].Code;
                            worksheet.Cells[(5 + i), 2].Value = objCustomerRttList[i].FullName;
                            if ((objCustomerRttList[i].AddressList != null) && (objCustomerRttList[i].AddressList.Count > 0))
                            {
                                worksheet.Cells[(5 + i), 3].Value = objCustomerRttList[i].AddressList[0].FullName;
                            }
                        }
                        worksheet.Cells[2, 4, (5 + objCustomerRttList.Count), 5].Style.Locked = false;
                        worksheet.Cells[4, 6].Value = "Для выбора адреса доставки, поставьте \"+\" в столбце \"Вкл\"";
                    }

                    worksheet.Cells["A1:F1000"].AutoFitColumns();
                    worksheet.Protection.SetPassword("A1");

                    //((Excel._Worksheet)oWB.Worksheets[1]).Activate();
                    //oSheet.Protect("A1", true, true, true, false, false, false, false, false, false, false, false, false, false, false, false);

                    worksheet = package.Workbook.Worksheets.Add("SupplItems");
                    worksheet.Cells[1, 1].Value = "Товарная марка";
                    worksheet.Cells[1, 2].Value = "Товарная группа";
                    worksheet.Cells[1, 3].Value = "Товарная подгруппа";
                    worksheet.Cells[1, 4].Value = "Код товара";
                    worksheet.Cells[1, 5].Value = "Наименование товара";
                    worksheet.Cells[1, 6].Value = "Артикул товара";
                    worksheet.Cells[1, 7].Value = "Цена, руб.";
                    worksheet.Cells[1, 8].Value = "Заказано";
                    worksheet.Cells[1, 9].Value = "Сумма, руб.";

                    if (bShowStockQty == true)
                    {
                        worksheet.Cells[1, 10].Value = "Остаток на складе";
                    }

                    using (var range = worksheet.Cells[1, 1, 1, 10])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Font.Size = 12;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.White);
                        range.Style.Font.Color.SetColor(Color.Black);
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                    }
                    worksheet.Row(1).Height = 20;

                    System.Int32 iLastIndxRowForPrint = 2;
                    if (objBlankItemList != null)
                    {
                        foreach (CBlankItem objBlankItem in objBlankItemList)
                        {
                            worksheet.Cells[iLastIndxRowForPrint, 1].Value = objBlankItem.ProductOwnerName;
                            worksheet.Cells[iLastIndxRowForPrint, 2].Value = objBlankItem.PartTypeName;
                            worksheet.Cells[iLastIndxRowForPrint, 3].Value = objBlankItem.PartSubTypeName;
                            worksheet.Cells[iLastIndxRowForPrint, 4].Value = objBlankItem.PartsId;
                            worksheet.Cells[iLastIndxRowForPrint, 5].Value = objBlankItem.PartsName;
                            worksheet.Cells[iLastIndxRowForPrint, 6].Value = objBlankItem.PartsArticle;
                            worksheet.Cells[iLastIndxRowForPrint, 7].Value = objBlankItem.Price;
                            worksheet.Cells[iLastIndxRowForPrint, 8].Value = 0;
                            worksheet.Cells[iLastIndxRowForPrint, 9, iLastIndxRowForPrint, 9].FormulaR1C1 = "=RC[-2]*RC[-1]";
                            if (bShowStockQty == true)
                            {
                                worksheet.Cells[iLastIndxRowForPrint, 10].Value = objBlankItem.CurrentQty;
                            }

                            iLastIndxRowForPrint++;
                        }
                    }
                    worksheet.Cells["A1:G1000"].AutoFitColumns();
                    worksheet.Cells[1, 1, 1, 9].AutoFilter = true;
                    worksheet.Cells[1, 8, 10000, 8].Style.Locked = false;
                    worksheet.Cells[1, 1, 1, 26].Style.Locked = false;

                    //worksheet.Cells["H1:H65000"].Style.Locked = false;
                    //worksheet.Cells["A1:Z1"].Style.Locked = false;

                    worksheet.Protection.SetPassword("A1");
                    worksheet.Protection.AllowAutoFilter = true;

                    worksheet = null;

                    package.Save();

                    try
                    {
                        using (System.Diagnostics.Process process = new System.Diagnostics.Process())
                        {
                            process.StartInfo.FileName = strFileName;
                            process.StartInfo.Verb = "Open";
                            process.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
                            process.Start();
                        }
                    }
                    catch
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show(this, "Не удалось найти приложение, чтобы открыть файл \"" + strFileName + "\"", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(
                    "ExportToExcelBlank.\nТекст ошибки: " + f.Message, "Ошибка",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            return;
        }
        
        private void ExportToExcel(ERP_Mercury.Common.CCustomer objCustomer, ERP_Mercury.Common.CChildDepart objChildCustomer, 
            ERP_Mercury.Common.CDepart objDepart,
            List<CBlankItem> objBlankItemList, List<ERP_Mercury.Common.CRtt> objCustomerRttList)
        {
            Excel.Application oXL = null;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            try
            {
                this.Cursor = Cursors.WaitCursor;
                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));

                oSheet = (Excel._Worksheet)oWB.Worksheets[1];

                //Клиент
                oSheet.Name = "Suppl";
                oSheet.Cells[1, 1] = "Код клиента";
                oSheet.Cells[1, 2] = "Клиент";
                oSheet.Cells[1, 3] = "Подразделение";
                oSheet.Cells[1, 4] = "Дочерний клиент";
                oSheet.Cells[1, 5] = "Бонус";
                oSheet.Cells[1, 6] = "Примечание";
                //oSheet.Cells[2, 5] = "Для помеки \"бонусный заказ\", поставьте \"+\" в ячейке E2";
                oSheet.get_Range(oSheet.Cells[2, 6], oSheet.Cells[2, 6]).Locked = false;
                if (objCustomer != null) 
                { 
                    oSheet.Cells[2, 1] = objCustomer.InterBaseID;
                    oSheet.Cells[2, 2] = objCustomer.FullName;
                }
                if (objDepart != null)
                {
                    oSheet.Cells[2, 3] = objDepart.DepartCode;
                }
                if (objChildCustomer != null)
                {
                    oSheet.Cells[2, 4] = objChildCustomer.Code;
                }
                for (System.Int32 i = 1; i <= 6; i++)
                {
                    oSheet.get_Range(oSheet.Cells[1, i], oSheet.Cells[1, i]).Font.Bold = true;
                    oSheet.get_Range(oSheet.Cells[1, i], oSheet.Cells[1, i]).Font.Size = 12;
                    oSheet.get_Range(oSheet.Cells[1, i], oSheet.Cells[1, i]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    oSheet.get_Range(oSheet.Cells[4, i], oSheet.Cells[1, i]).Font.Bold = true;
                    oSheet.get_Range(oSheet.Cells[4, i], oSheet.Cells[1, i]).Font.Size = 12;
                    oSheet.get_Range(oSheet.Cells[4, i], oSheet.Cells[1, i]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }

                // попробуем воткнуть розничные точки
                if ((objCustomerRttList != null) && (objCustomerRttList.Count > 0))
                {
                    oSheet.Cells[4, 1] = "Код РТТ";
                    oSheet.Cells[4, 2] = "Наименование РТТ";
                    oSheet.Cells[4, 3] = "Адрес";
                    oSheet.Cells[4, 4] = "Вкл";

                    for (System.Int32 i = 0; i < objCustomerRttList.Count; i++)
                    {
                        oSheet.Cells[(5 + i), 1] = objCustomerRttList[ i ].Code;
                        oSheet.Cells[(5 + i), 2] = objCustomerRttList[ i ].FullName;
                        if ( ( objCustomerRttList[i].AddressList != null ) && ( objCustomerRttList[ i ].AddressList.Count > 0 ) )
                        {
                            oSheet.Cells[(5 + i), 3] = objCustomerRttList[ i ].AddressList[0].FullName;
                        }
                    }
                    oSheet.get_Range(oSheet.Cells[2, 4], oSheet.Cells[(5 + objCustomerRttList.Count), 5]).Locked = false;
                    oSheet.Cells[4, 6] = "Для выбора адреса доставки, поставьте \"+\" в столбце \"Вкл\"";
                }

                oSheet.get_Range("A1", "F1").EntireColumn.AutoFit();

                ((Excel._Worksheet)oWB.Worksheets[1]).Activate();
                oSheet.Protect("A1", true, true, true, false, false, false, false, false, false, false, false, false, false, false, false);

                oSheet = (Excel._Worksheet)oWB.Worksheets[ 2 ];
                oSheet.Name = "SupplItems";
                oSheet.Cells[1, 1] = "Товарная марка";
                oSheet.Cells[1, 2] = "Товарная группа";
                oSheet.Cells[1, 3] = "Товарная подгруппа";
                oSheet.Cells[1, 4] = "Код товара";
                oSheet.Cells[1, 5] = "Наименование товара";
                oSheet.Cells[1, 6] = "Артикул товара";
                oSheet.Cells[1, 7] = "Цена, руб.";
                oSheet.Cells[1, 8] = "Заказано";
                oSheet.Cells[1, 9] = "Сумма, руб.";
                for (System.Int32 i = 1; i <= 9; i++)
                {
                    oSheet.get_Range(oSheet.Cells[1, i], oSheet.Cells[1, i]).Font.Bold = true;
                    oSheet.get_Range(oSheet.Cells[1, i], oSheet.Cells[1, i]).Font.Size = 12;
                    oSheet.get_Range(oSheet.Cells[1, i], oSheet.Cells[1, i]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
                System.Int32 iLastIndxRowForPrint = 2;
                if (objBlankItemList != null)
                {
                    foreach (CBlankItem objBlankItem in objBlankItemList)
                    {
                        oSheet.Cells[iLastIndxRowForPrint, 1] = objBlankItem.ProductOwnerName;
                        oSheet.Cells[iLastIndxRowForPrint, 2] = objBlankItem.PartTypeName;
                        oSheet.Cells[iLastIndxRowForPrint, 3] = objBlankItem.PartSubTypeName;
                        oSheet.Cells[iLastIndxRowForPrint, 4] = objBlankItem.PartsId;
                        oSheet.Cells[iLastIndxRowForPrint, 5] = objBlankItem.PartsName;
                        oSheet.Cells[iLastIndxRowForPrint, 6] = objBlankItem.PartsArticle;
                        oSheet.Cells[iLastIndxRowForPrint, 7] = objBlankItem.Price;
                        oSheet.Cells[iLastIndxRowForPrint, 8] = 0;
                        oSheet.get_Range(oSheet.Cells[iLastIndxRowForPrint, 9], oSheet.Cells[iLastIndxRowForPrint, 9]).Formula = "=RC[-2]*RC[-1]";
                        iLastIndxRowForPrint++;
                    }
                }
                oSheet.get_Range("A1", "G1").EntireColumn.AutoFit();

                oSheet.get_Range("A1", "A1").AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
                oSheet.get_Range("H1", "H65000").Locked = false;
                oSheet.get_Range("A1", "Z1").Locked = false;

                oSheet.Protect("A1", true, true, true, true, true, true, true, false, false, false, false, false, false, true, false);
                //oSheet.Protect("A1", false, true, false, false, false, false, false, true, true, true, true, true, false, false, true);



                oXL.Visible = true;
                oXL.UserControl = true;
            }
            catch (System.Exception f)
            {
                if (oXL != null) { oXL.Quit(); }
                oXL = null;
                DevExpress.XtraEditors.XtraMessageBox.Show(
                    "Ошибка экспорта в MS Excel.\n\nТекст ошибки: " + f.Message, "Ошибка",
                   System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = System.Windows.Forms.Cursors.Default;
            }
        }
        private void btnCreateBlank_Click(object sender, EventArgs e)
        {
            try
            {

                Cursor.Current = Cursors.WaitCursor;
                ERP_Mercury.Common.CCustomer objCustomer = (cboxCustomer.SelectedItem == null) ? null : (ERP_Mercury.Common.CCustomer)cboxCustomer.SelectedItem;
                ERP_Mercury.Common.CDepart objDepart = (cboxDepart.SelectedItem == null) ? null : (ERP_Mercury.Common.CDepart)cboxDepart.SelectedItem;
                ERP_Mercury.Common.CChildDepart objChildCustomer = (cboxChildCust.SelectedItem == null) ? null : (ERP_Mercury.Common.CChildDepart)cboxChildCust.SelectedItem;
                ERP_Mercury.Common.CDepartTeam objDepartTeam = (cboxDepartTeam.SelectedItem == null) ? null : (ERP_Mercury.Common.CDepartTeam)cboxDepartTeam.SelectedItem;
                List<ERP_Mercury.Common.CRtt> objCustomerRttList = ERP_Mercury.Common.CRtt.GetRttList(m_objProfile, null, objCustomer.ID);
                
                List<CBlankItem> objBlankItemListFull = CBlankItem.GetBlankItemList(m_objProfile, objDepartTeam.uuidID, checkStockQty.Checked);
                List<CBlankItem> objBlankItemList = new List<CBlankItem>();
                if ((objBlankItemListFull != null) && (checklstboxProductOwner.CheckedItems.Count > 0))
                {
                    foreach (CBlankItem objBlankItem in objBlankItemListFull)
                    {
                        for (System.Int32 i = 0; i < checklstboxProductOwner.Items.Count; i++)
                        {
                            if (checklstboxProductOwner.Items[i].CheckState == CheckState.Checked)
                            {
                                if (objBlankItem.ProductOwnerName == ((ERP_Mercury.Common.CProductOwner)checklstboxProductOwner.Items[i].Value).Name)
                                {
                                    objBlankItemList.Add(objBlankItem);
                                    break;
                                }
                            }
                        }
                    }
                }

                //ExportToExcel(objCustomer, objChildCustomer, objDepart, objBlankItemList, objCustomerRttList);
                ExportToExcelBlank(objCustomer, objChildCustomer, objDepart, objBlankItemList, objCustomerRttList, checkStockQty.Checked);
                Cursor.Current = Cursors.Default;

            }
            catch (System.Exception f)
            {
                System.Windows.Forms.MessageBox.Show(this, "Ошибка печати\n" + f.Message, "Ошибка",
                   System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            return;

        }
        #endregion

        #region Открыть файл
        /// <summary>
        /// Очистка содержимого закладки "Заказ"
        /// </summary>
        private void ClearImportTab()
        {
            try
            {
                this.tableLayoutPanelSupplBgrnd.SuspendLayout();
                ((System.ComponentModel.ISupportInitialize)(this.treeListSupplItms)).BeginInit();

                txtCustomer.Text = "";
                txtDepartCode.Text = "";
                txtChildCustCode.Text = "";
                txtRtt.Text = "";
                checkBonus.Checked = false;
                dtDateDelivery.DateTime = System.DateTime.Today;

                treeListSupplItms.Nodes.Clear();

                btnOpenFile.Enabled = true;

                this.tableLayoutPanelSupplBgrnd.ResumeLayout(false);
                ((System.ComponentModel.ISupportInitialize)(this.treeListSupplItms)).EndInit();

            }
            catch (System.Exception f)
            {
                SendMessageToLog("ClearImportTab. Текст ошибки: " + f.Message);
            }
            finally
            {
            }
            return;
        }

        private void OpenFile()
        {
            try
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    this.Refresh();
                    if( (openFileDialog.FileName != "") && ( System.IO.File.Exists( openFileDialog.FileName ) == true ) )
                    {
                        Cursor = Cursors.WaitCursor;
                        ClearImportTab();
                        SendMessageToLog("Идет импорт данных из файла...");
                        this.Refresh();
                        ImportDataFromExcel(openFileDialog.FileName);
                        btnSaveToDB.Enabled = ( treeListSupplItms.Nodes.Count > 0 );
                        Cursor = Cursors.Default;
                    }
                }

            }
            catch (System.Exception f)
            {
                SendMessageToLog("OpenFile. Текст ошибки: " + f.Message);
            }
            finally
            {
            }
            return;
        }
        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFile();

            }
            catch (System.Exception f)
            {
                SendMessageToLog("btnOpenFile_Click. Текст ошибки: " + f.Message);
            }
            finally
            {
            }
            return;
        }
        /// <summary>
        /// Импорт данных из MS Excel
        /// </summary>
        /// <param name="strFileName">файл MS Excel
        /// </param>
        private void ImportDataFromExcel( System.String strFileName )
        {
            Excel.Application oXL = null;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            this.tableLayoutPanelSupplBgrnd.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.treeListSupplItms)).BeginInit();

            try
            {
                this.Cursor = Cursors.WaitCursor;
                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open( strFileName, 0, true, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value));

                if ((oWB.Worksheets.Count >= 2) && (((Excel._Worksheet)oWB.Worksheets[1]).Name == "Suppl") && (((Excel._Worksheet)oWB.Worksheets[2]).Name == "SupplItems"))
                {
                    oSheet = (Excel._Worksheet)oWB.Worksheets[1];
                    if ((System.Convert.ToString(oSheet.get_Range("A1", Missing.Value).Value2).CompareTo("Код клиента") == 0) &&
                        (System.Convert.ToString(oSheet.get_Range("A2", Missing.Value).Value2).CompareTo("") != 0) &&
                        (System.Convert.ToString(oSheet.get_Range("C2", Missing.Value).Value2).CompareTo("") != 0))
                    {
                        txtCustomer.Text = System.Convert.ToString( oSheet.get_Range("B2", Missing.Value).Value2 );
                        txtCustomer.Tag = System.Convert.ToInt32(oSheet.get_Range("A2", Missing.Value).Value2);
                        txtDepartCode.Text = System.Convert.ToString(oSheet.get_Range("C2", Missing.Value).Value2);
                        txtChildCustCode.Text = System.Convert.ToString(oSheet.get_Range("D2", Missing.Value).Value2);
                        checkBonus.Checked = (System.Convert.ToString(oSheet.get_Range("E2", Missing.Value).Value2).Length > 0);
                        memoDescrpn.Text = System.Convert.ToString(oSheet.get_Range("F2", Missing.Value).Value2);
                    }

                    // проверка на РТТ
                    if( System.Convert.ToString(oSheet.get_Range("A5", Missing.Value).Value2).CompareTo("") != 0 )
                    {
                        // у нас есть хотя бы одна РТТ
                        // нужно понять на какой стоит "+"
                        System.Int32 iRttCount = 1;
                        while (System.Convert.ToString(oSheet.get_Range(oSheet.Cells[5 + iRttCount, 1], oSheet.Cells[5 + iRttCount, 1]).Value2).CompareTo("") != 0)
                        {
                            iRttCount++;
                        }
                        // мы сосчитали количество РТТ, теперь нужно понять, где стоит "+"
                        for( System.Int32 i = 0; i < iRttCount; i++ )
                        {
                            if(System.Convert.ToString(oSheet.get_Range(oSheet.Cells[5 + i, 4], oSheet.Cells[5 + i, 4]).Value2).CompareTo("+") == 0)
                            {
                                txtRtt.Text = System.Convert.ToString(oSheet.get_Range(oSheet.Cells[5 + i, 1], oSheet.Cells[5 + i, 1]).Value2);
                                break;
                            }
                        }

                        txtCustomer.Text = System.Convert.ToString(oSheet.get_Range("B2", Missing.Value).Value2);
                        txtCustomer.Tag = System.Convert.ToInt32(oSheet.get_Range("A2", Missing.Value).Value2);
                        txtDepartCode.Text = System.Convert.ToString(oSheet.get_Range("C2", Missing.Value).Value2);
                        txtChildCustCode.Text = System.Convert.ToString(oSheet.get_Range("D2", Missing.Value).Value2);
                    }


                    oSheet = (Excel._Worksheet)oWB.Worksheets[2];
                    System.Boolean bEndList = false;
                    System.Int32 iRowIndex = 2;
                    System.Int32 iPartsId = 0;
                    System.Int32 iOrderQty = 0;
                    System.String strPartsName = "";
                    while (bEndList == false)
                    {
                        iPartsId = 0;
                        iOrderQty = 0;
                        strPartsName = "";
                        try
                        {
                            strPartsName = System.Convert.ToString(oSheet.get_Range(oSheet.Cells[iRowIndex, 5], oSheet.Cells[iRowIndex, 5]).Value2);
                            iPartsId = System.Convert.ToInt32(oSheet.get_Range(oSheet.Cells[iRowIndex, 4], oSheet.Cells[iRowIndex, 4]).Value2);
                            iOrderQty = System.Convert.ToInt32(oSheet.get_Range(oSheet.Cells[iRowIndex, 8], oSheet.Cells[iRowIndex, 8]).Value2);
                        }
                        catch
                        {
                            iPartsId = 0;
                            iOrderQty = 0;
                            strPartsName = "";
                        }
                        if( ( strPartsName != "" ) && ( iPartsId != 0 ) && ( iOrderQty != 0 ) )
                        {
                            treeListSupplItms.AppendNode(new object[] { strPartsName, iOrderQty }, null).Tag = iPartsId;
                        }
                        iRowIndex++;
                        if (System.Convert.ToString(oSheet.get_Range(oSheet.Cells[iRowIndex, 1], oSheet.Cells[iRowIndex, 1]).Value2) == "")
                        {
                            bEndList = true;
                        }
                    }

                }

                oSheet = null;
                oWB = null;
                oXL.Quit();
                oXL = null;
            }
            catch (System.Exception f)
            {
                if (oXL != null) { oXL.Quit(); }
                oXL = null;
                DevExpress.XtraEditors.XtraMessageBox.Show(
                    "Ошибка импорта данных из MS Excel.\n\nТекст ошибки: " + f.Message, "Ошибка",
                   System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                this.tableLayoutPanelSupplBgrnd.ResumeLayout(false);
                ((System.ComponentModel.ISupportInitialize)(this.treeListSupplItms)).EndInit();
                this.Cursor = System.Windows.Forms.Cursors.Default;
            }
        }
        #endregion

        #region Запись заказа в БД
        /// <summary>
        /// Запись заказа в БД
        /// </summary>
        private void SendSupplToDB()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                pictureEdit1.Image = ERPMercuryImportSuppl.Properties.Resources.note_small;
                System.Guid SupplId = System.Guid.Empty;
                if ((txtCustomer.Text != "") && (txtCustomer.Tag != null) && (txtDepartCode.Text != "") && (treeListSupplItms.Nodes.Count > 0))
                {
                    m_dsOrders.Tables["tblSuppl"].Rows.Clear();
                    m_dsOrders.Tables["tblSupplItems"].Rows.Clear();

                    System.Data.SqlClient.SqlConnection DBConnection = m_objProfile.GetDBSource();
                    if (DBConnection == null)
                    {
                        SendMessageToLog("Отсутствует соединение с БД.");
                        return;
                    }
                    SendMessageToLog("Идет сохранение заказа в БД...");
                    System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
                    cmd.Connection = DBConnection;
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.CommandText = System.String.Format("[{0}].[dbo].[usp_ConvertSupplFromExcel]", m_objProfile.GetOptionsDllDBName());
                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@CustomerId", System.Data.SqlDbType.Int));
                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@DepartCode", System.Data.SqlDbType.NVarChar, 3));

                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Suppl_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Depart_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Customer_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@CustomerChild_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Rtt_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Address_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ERROR_NUM", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ERROR_MES", System.Data.SqlDbType.NVarChar, 4000));
                    cmd.Parameters["@ERROR_MES"].Direction = System.Data.ParameterDirection.Output;


                    cmd.Parameters["@CustomerId"].Value = System.Convert.ToInt32( txtCustomer.Tag );
                    cmd.Parameters["@DepartCode"].Value = txtDepartCode.Text;

                    if (txtChildCustCode.Text != "")
                    {
                        cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ChildDepartCode", System.Data.SqlDbType.NVarChar, 56));
                        cmd.Parameters["@ChildDepartCode"].Value = txtChildCustCode.Text;
                    }
                    if (txtRtt.Text != "")
                    {
                        cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@RttCode", System.Data.SqlDbType.NVarChar, 56));
                        cmd.Parameters["@RttCode"].Value = txtRtt.Text;
                    }
                    cmd.ExecuteNonQuery();
                    System.Int32 iRes = (System.Int32)cmd.Parameters["@RETURN_VALUE"].Value;
                    if (iRes == 0)
                    {
                        if ((cmd.Parameters["@Depart_Guid"].Value != System.DBNull.Value) &&
                            (cmd.Parameters["@Customer_Guid"].Value != System.DBNull.Value))
                        {
                            System.Data.DataRow newRow = m_dsOrders.Tables["tblSuppl"].NewRow();
                            SupplId = (System.Guid)cmd.Parameters["@Suppl_Guid"].Value;
                            newRow["Suppl_Guid"] = SupplId;
                            newRow["Customer_Guid"] = (System.Guid)cmd.Parameters["@Customer_Guid"].Value;
                            newRow["Depart_Guid"] = (System.Guid)cmd.Parameters["@Depart_Guid"].Value;
                            newRow["Suppl_Num"] = 0;
                            newRow["Suppl_BeginDate"] = System.DateTime.Today;
                            newRow["Suppl_State"] = iSupplState;
                            newRow["Suppl_Bonus"] = ( ( checkBonus.Checked == true ) ? 1 : 0 );
                            newRow["Suppl_Note"] = (memoDescrpn.Text == "") ? "-" : memoDescrpn.Text;
                            newRow["SupplType_Guid"] = System.DBNull.Value;
                            newRow["Suppl_Version"] = System.DBNull.Value;
                            newRow["Suppl_DeliveryDate"] = dtDateDelivery.DateTime;

                            if ((cmd.Parameters["@Rtt_Guid"].Value != System.DBNull.Value) &&
                                (cmd.Parameters["@Address_Guid"].Value != System.DBNull.Value))
                            {

                                newRow["Rtt_Guid"] = (System.Guid)cmd.Parameters["@Rtt_Guid"].Value;
                                newRow["Address_Guid"] = (System.Guid)cmd.Parameters["@Address_Guid"].Value;
                            }
                            if (txtChildCustCode.Text != "" )
                            {
                                 newRow["SupplType_Guid"] = new System.Guid( "9FD32373-ABB4-4C6C-8225-B3BF88F22AB0" );
                            }
                            m_dsOrders.Tables["tblSuppl"].Rows.Add(newRow);

                            // теперь обработаем содержимое заказа
                            cmd.Parameters.Clear();
                            cmd.CommandText = System.String.Format("[{0}].[dbo].[usp_ConvertSupplItmsFromExcel]", m_objProfile.GetOptionsDllDBName());
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
                            System.Data.DataRow newRowItem = null;

                            foreach (DevExpress.XtraTreeList.Nodes.TreeListNode objNode in treeListSupplItms.Nodes)
                            {
                                if (objNode.Tag == null) { continue; }
                                cmd.Parameters["@PartsId"].Value = System.Convert.ToInt32(objNode.Tag);
                                cmd.Parameters["@PartsQty"].Value = System.Convert.ToInt32(objNode.GetValue(colOrderQty));
                                cmd.ExecuteNonQuery();
                                iRes = (System.Int32)cmd.Parameters["@RETURN_VALUE"].Value;
                                if (iRes == 0)
                                {
                                    newRowItem = m_dsOrders.Tables["tblSupplItems"].NewRow();
                                    newRowItem["SupplItem_Guid"] = (System.Guid)cmd.Parameters["@SupplItms_Guid"].Value;
                                    newRowItem["Suppl_Guid"] = SupplId;
                                    newRowItem["Parts_Guid"] = (System.Guid)cmd.Parameters["@Parts_Guid"].Value;
                                    newRowItem["Measure_Guid"] = (System.Guid)cmd.Parameters["@Measure_Guid"].Value;
                                    newRowItem["SupplItem_OrderQuantity"] = System.Convert.ToInt32(cmd.Parameters["@SplItms_OrderQty"].Value);
                                    newRowItem["SupplItem_Quantity"] = System.Convert.ToInt32(cmd.Parameters["@SplItms_Quatity"].Value);

                                    m_dsOrders.Tables["tblSupplItems"].Rows.Add(newRowItem);
                                }
                                else
                                {
                                    SendMessageToLog((System.String)objNode.GetValue(colPartaName) + " " + (System.String)cmd.Parameters["@ERROR_MES"].Value);
                                }
                            }

                            m_dsOrders.AcceptChanges();         
               
                            //теперь попробуем воспользоваться web сервисом
                            cmd.Parameters.Clear();
                            cmd.CommandText = System.String.Format("[{0}].[dbo].[usp_GetWebServiceName]", m_objProfile.GetOptionsDllDBName());
                            cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                            cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@SoapUrl", System.Data.SqlDbType.NVarChar, 400));
                            cmd.Parameters["@SoapUrl"].Direction = System.Data.ParameterDirection.Output;
                            cmd.ExecuteNonQuery();
                            iRes = (System.Int32)cmd.Parameters["@RETURN_VALUE"].Value;
                            System.Int32 iRetCode = -1;

                            if (iRes == 0)
                            {
                                SendMessageToLog("web сервис: " + (System.String)cmd.Parameters["@SoapUrl"].Value);
                                SendMessageToLog("идет вызов web сервиса");
                                ERPMercuryImportSuppl.WebReference.SOAPService objSrvGprs = null;
                                try
                                {
                                    SendMessageToLog("v-iis01.SOAPService objSrvGprs = new SalesManager.v-iis01.SOAPService()");
                                    objSrvGprs = new ERPMercuryImportSuppl.WebReference.SOAPService(); // .srvGPRS();

                                    SendMessageToLog("объект \"сервис\" создан");

                                    objSrvGprs.Url = (System.String)cmd.Parameters["@SoapUrl"].Value;
                                    
                                    SendMessageToLog("сервису присвоен адрес: " + objSrvGprs.Url);

                                    SendMessageToLog("вызывается метод objSrvGprs.SaveNewSuppl");

                                    SendMessageToLog("количество таблиц в наборе" + m_dsOrders.Tables.Count.ToString());
                                    SendMessageToLog("количество записей в таблице tblSuppl: " + m_dsOrders.Tables["tblSuppl"].Rows.Count.ToString());
                                    SendMessageToLog("количество записей в таблице tblSupplItems: " + m_dsOrders.Tables["tblSupplItems"].Rows.Count.ToString());

                                    iRetCode = objSrvGprs.SaveNewSuppl(m_dsOrders);

                                    SendMessageToLog("вызов метода objSrvGprs.SaveNewSuppl завершен");
                                }
                                catch (System.Exception ferr)
                                {
                                    DevExpress.XtraEditors.XtraMessageBox.Show(this, "вызов сервиса. Текст ошибки: " + ferr.Message, "Информация",
                                      System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);

                                    SendMessageToLog("вызов сервиса. Текст ошибки: " + ferr.Message);
                                }

                                //objSrvGprs.Url = "http://192.168.7.27/MercuryPDA/SOAPService.asmx";

                                if (iRetCode == 0)
                                {
                                    SendMessageToLog("заказ успешно обработан.");
                                    DevExpress.XtraEditors.XtraMessageBox.Show(this, "Заказ успешно обработан.", "Информация",
                                      System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                                    pictureEdit1.Image = ERPMercuryImportSuppl.Properties.Resources.ok_16;
                                }
                                else
                                {
                                    SendMessageToLog("ошибка передачи информации в БД. Код ошибки: " + iRetCode.ToString());
                                    DevExpress.XtraEditors.XtraMessageBox.Show(this, "Ошибка передачи информации в БД.", "Ошибка",
                                      System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                                    pictureEdit1.Image = ERPMercuryImportSuppl.Properties.Resources.warning;
                                }
                                objSrvGprs = null;
                            }
                            else
                            {
                                SendMessageToLog("ошибка передачи информации в БД.");
                                DevExpress.XtraEditors.XtraMessageBox.Show(this, "Ошибка передачи информации в БД. Не удалось получить адрес web сервиса.", "Ошибка",
                                  System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                                pictureEdit1.Image = ERPMercuryImportSuppl.Properties.Resources.warning;
                            }
                        }
                    }
                    else
                    {
                        SendMessageToLog(( System.String )cmd.Parameters["@ERROR_MES"].Value);
                    }

                    cmd = null;
                    DBConnection.Close();
                }
            }
            catch (System.Exception f)
            {
                SendMessageToLog("SendSupplToDB. Текст ошибки: " + f.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
            return;
        }

        private void btnSaveToDB_Click(object sender, EventArgs e)
        {
            try
            {
                SendSupplToDB();
            }
            catch (System.Exception f)
            {
                SendMessageToLog("btnSaveToDB_Click. Текст ошибки: " + f.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
            return;
        }
        #endregion

        #region Импорт заказа из файла со списком штрих-кодов и количеством

        #region Выпадающий списки
        /// <summary>
        /// Стартует поток, в котором загружается список подразделений
        /// </summary>
        public void StartThreadLoadDepartList()
        {
            try
            {
                // инициализируем делегаты
                m_LoadDepartListDelegate = new LoadDepartListDelegate(LoadDepartListInComboBox);

                cboxImportSupplRtt.Text = strWaitLoadList;
                cboxImportSupplChildDepart.Text = strWaitLoadList;
                cboxImportSupplCustomer.Text = strWaitLoadList;
                cboxImportSupplDepart.Text = strWaitLoadList;

                cboxImportSupplRtt.Properties.Items.Clear();
                cboxImportSupplChildDepart.Properties.Items.Clear();
                cboxImportSupplCustomer.Properties.Items.Clear();
                cboxImportSupplDepart.Properties.Items.Clear();

                // запуск потока
                this.ThreadLoadDepartList = new System.Threading.Thread(LoadDepartListInThread);
                this.ThreadLoadDepartList.Start();
            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("StartThreadLoadDepartList().\n\nТекст ошибки: " + f.Message, "Ошибка",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            return;
        }
        /// <summary>
        /// Загружает список подразделений (метод, выполняемый в потоке)
        /// </summary>
        public void LoadDepartListInThread()
        {
            try
            {
                List<ERP_Mercury.Common.CDepart> objDepartList = ERP_Mercury.Common.CDepart.GetDepartList(m_objProfile, true );

                List<ERP_Mercury.Common.CDepart> objAddDepartList = new List<ERP_Mercury.Common.CDepart>();
                if ((objDepartList != null) && (objDepartList.Count > 0))
                {
                    System.Int32 iRecCount = 0;
                    System.Int32 iRecAllCount = 0;
                    foreach (ERP_Mercury.Common.CDepart objDepart in objDepartList)
                    {
                        objAddDepartList.Add(objDepart);
                        iRecCount++;
                        iRecAllCount++;

                        if (iRecCount == iRowsPartForLoadInComboBox)
                        {
                            iRecCount = 0;
                            Thread.Sleep(iThreadSleepTime);
                            this.Invoke(m_LoadDepartListDelegate, new Object[] { objAddDepartList, iRecAllCount });
                            objAddDepartList.Clear();
                        }

                    }
                    if (iRecCount != iRowsPartForLoadInComboBox)
                    {
                        iRecCount = 0;
                        this.Invoke(m_LoadDepartListDelegate, new Object[] { objAddDepartList, iRecAllCount });
                        objAddDepartList.Clear();
                    }

                }

                objDepartList = null;
                objAddDepartList = null;
                this.Invoke(m_LoadDepartListDelegate, new Object[] { null, 0 });
            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("LoadDepartListInThread.\n\nТекст ошибки: " + f.Message, "Ошибка",
                   System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
            }
            return;
        }
        /// <summary>
        /// Загрузка в выпадающий список с подразделениями порции значений
        /// </summary>
        /// <param name="objDepartList">порция из списка подразделений</param>
        /// <param name="iRowCountInList">всего записей в списке подразделений</param>
        private void LoadDepartListInComboBox(List<ERP_Mercury.Common.CDepart> objDepartList, System.Int32 iRowCountInList)
        {
            try
            {
                if ((objDepartList != null) && (objDepartList.Count > 0) && ( cboxImportSupplDepart.Properties.Items.Count < iRowCountInList))
                {
                    cboxImportSupplDepart.Properties.Items.AddRange(objDepartList);
                }
                else
                {
                    // процесс загрузки данных завершён
                    Thread.Sleep(iThreadSleepTime);

                    cboxImportSupplRtt.Text = "";
                    cboxImportSupplChildDepart.Text = "";
                    cboxImportSupplCustomer.Text = "";
                    cboxImportSupplDepart.Text = "";

                    cboxImportSupplRtt.SelectedItem = null;
                    cboxImportSupplChildDepart.SelectedItem = null;
                    cboxImportSupplCustomer.SelectedItem = null;
                    cboxImportSupplDepart.SelectedItem = null;

                    ValidatePropertiesBeforeImportInDB();
                }

            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("LoadDepartListInComboBox.\n\nТекст ошибки: " + f.Message, "Ошибка",
                   System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
            }
            return;
        }

        /// <summary>
        /// Обработчик "Выбрано Подразделение"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboxImportSupplDepart_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                cboxImportSupplChildDepart.SelectedItem = null;
                cboxImportSupplCustomer.SelectedItem = null;
                cboxImportSupplRtt.SelectedItem = null;

                cboxImportSupplChildDepart.Properties.Items.Clear();
                cboxImportSupplCustomer.Properties.Items.Clear();
                cboxImportSupplRtt.Properties.Items.Clear();

                if (cboxImportSupplDepart.SelectedItem == null) { return; }

                ERP_Mercury.Common.CDepart objDepart = (ERP_Mercury.Common.CDepart)cboxImportSupplDepart.SelectedItem;

                if ((objDepart.CustomerList != null) && (objDepart.CustomerList.Count > 0))
                {
                    cboxImportSupplCustomer.Properties.Items.AddRange(objDepart.CustomerList);
                }

                if ((cboxImportSupplDepart.Properties.Items.Count > 0) && (cboxImportSupplDepart.SelectedItem != null))
                {
                    cboxImportSupplCustomer.Focus();
                }
            }
            catch (System.Exception f)
            {
                SendMessageToLog("cboxImportSupplDepart_SelectedValueChanged. Текст ошибки: " + f.Message);
            }
            finally
            {
                ValidatePropertiesBeforeImportInDB();

                Cursor = Cursors.Default;
            }
            return;
        }
        /// <summary>
        /// Обработчик "Выбран Клиент"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboxImportSupplCustomer_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {

                cboxImportSupplChildDepart.SelectedItem = null;
                cboxImportSupplChildDepart.Properties.Items.Clear();
                cboxImportSupplRtt.SelectedItem = null;
                cboxImportSupplRtt.Properties.Items.Clear();

                if (cboxImportSupplCustomer.SelectedItem == null) { return; }

                ERP_Mercury.Common.CCustomer objCustomer = (ERP_Mercury.Common.CCustomer)cboxImportSupplCustomer.SelectedItem;

                // дочерние подразделения
                objCustomer.ChildDepartList = ERP_Mercury.Common.CChildDepart.GetChildDepartList(m_objProfile, null, objCustomer.ID );
                if ((objCustomer.ChildDepartList != null) && (objCustomer.ChildDepartList.Count > 0))
                {
                    cboxImportSupplChildDepart.Properties.Items.AddRange(objCustomer.ChildDepartList);
                }

                // розничные точки
                List<ERP_Mercury.Common.CRtt> objCustomerRttList = ERP_Mercury.Common.CRtt.GetRttListForImportSuppl(m_objProfile, null, objCustomer.ID);
                if ((objCustomerRttList != null) && (objCustomerRttList.Count > 0))
                {
                    cboxImportSupplRtt.Properties.Items.AddRange(objCustomerRttList);
                }
                if (cboxImportSupplRtt.Properties.Items.Count > 0)
                {
                    cboxImportSupplRtt.SelectedItem = cboxImportSupplRtt.Properties.Items[0];
                }

                if ((cboxImportSupplCustomer.Properties.Items.Count > 0) && (cboxImportSupplCustomer.SelectedItem != null) )
                {
                    if (cboxImportSupplChildDepart.Properties.Items.Count > 0)
                    {
                        cboxImportSupplChildDepart.Focus();
                    }
                    else if (cboxImportSupplRtt.Properties.Items.Count > 0)
                    {
                        cboxImportSupplRtt.Focus();
                    }
                }

            }
            catch (System.Exception f)
            {
                SendMessageToLog("cboxImportSupplCustomer_SelectedValueChanged. Текст ошибки: " + f.Message);
            }
            finally
            {
                ValidatePropertiesBeforeImportInDB();

                Cursor = Cursors.Default;
            }
            return;
        }
        #endregion

        #region Запись заказа в БД
        /// <summary>
        /// Запись заказа в БД
        /// </summary>
        /// <param name="CustomerId">УИ клиента (InterBase)</param>
        /// <param name="DepartCode">код подразделения</param>
        /// <param name="ChildDepartCode">код дочернего клиента</param>
        /// <param name="RttCode">код РТТ</param>
        /// <param name="Suppl_Note">примечание к заказу</param>
        /// <param name="Suppl_Bonus">признак "Бонус"</param>
        /// <param name="Suppl_DeliveryDate">дата доставки</param>
        private void SendSupplToDB(System.Int32 CustomerId, System.String DepartCode, System.String ChildDepartCode, System.String RttCode,
            System.String Suppl_Note, System.Boolean Suppl_Bonus, System.DateTime Suppl_DeliveryDate)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                System.Guid SupplId = System.Guid.Empty;

                m_dsOrders.Tables["tblSuppl"].Rows.Clear();
                m_dsOrders.Tables["tblSupplItems"].Rows.Clear();

                System.Data.SqlClient.SqlConnection DBConnection = m_objProfile.GetDBSource();
                if (DBConnection == null)
                {
                    SendMessageToLog("Отсутствует соединение с БД.");
                    return;
                }
                SendMessageToLog("Идет сохранение заказа в БД...");
                System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
                cmd.Connection = DBConnection;
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandText = System.String.Format("[{0}].[dbo].[sp_ConvertSupplFromExcel]", m_objProfile.GetOptionsDllDBName());
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@CustomerId", System.Data.SqlDbType.Int));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@DepartCode", System.Data.SqlDbType.NVarChar, 3));

                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Suppl_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Depart_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Customer_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@CustomerChild_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Rtt_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Address_Guid", System.Data.SqlDbType.UniqueIdentifier, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ERROR_NUM", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ERROR_MES", System.Data.SqlDbType.NVarChar, 4000));
                cmd.Parameters["@ERROR_MES"].Direction = System.Data.ParameterDirection.Output;


                cmd.Parameters["@CustomerId"].Value = CustomerId;
                cmd.Parameters["@DepartCode"].Value = DepartCode;

                if (ChildDepartCode != "")
                {
                    cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ChildDepartCode", System.Data.SqlDbType.NVarChar, 56));
                    cmd.Parameters["@ChildDepartCode"].Value = ChildDepartCode;
                }
                if (RttCode != "")
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
                        System.Data.DataRow newRow = m_dsOrders.Tables["tblSuppl"].NewRow();
                        SupplId = (System.Guid)cmd.Parameters["@Suppl_Guid"].Value;
                        newRow["Suppl_Guid"] = SupplId;
                        newRow["Customer_Guid"] = (System.Guid)cmd.Parameters["@Customer_Guid"].Value;
                        newRow["Depart_Guid"] = (System.Guid)cmd.Parameters["@Depart_Guid"].Value;
                        newRow["Suppl_Num"] = 0;
                        newRow["Suppl_BeginDate"] = System.DateTime.Today;
                        newRow["Suppl_State"] = iSupplState;
                        newRow["Suppl_Bonus"] = ((Suppl_Bonus == true) ? 1 : 0);
                        newRow["Suppl_Note"] = Suppl_Note;
                        newRow["SupplType_Guid"] = System.DBNull.Value;
                        newRow["Suppl_Version"] = System.DBNull.Value;
                        newRow["Suppl_DeliveryDate"] = Suppl_DeliveryDate;

                        if ((cmd.Parameters["@Rtt_Guid"].Value != System.DBNull.Value) &&
                            (cmd.Parameters["@Address_Guid"].Value != System.DBNull.Value))
                        {

                            newRow["Rtt_Guid"] = (System.Guid)cmd.Parameters["@Rtt_Guid"].Value;
                            newRow["Address_Guid"] = (System.Guid)cmd.Parameters["@Address_Guid"].Value;
                        }
                        if (txtChildCustCode.Text != "")
                        {
                            newRow["SupplType_Guid"] = new System.Guid("9FD32373-ABB4-4C6C-8225-B3BF88F22AB0");
                        }
                        m_dsOrders.Tables["tblSuppl"].Rows.Add(newRow);

                        // теперь обработаем содержимое заказа
                        cmd.Parameters.Clear();
                        cmd.CommandText = System.String.Format("[{0}].[dbo].[sp_ConvertSupplItmsFromExcel]", m_objProfile.GetOptionsDllDBName());
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
                        System.Data.DataRow newRowItem = null;

                        foreach (DevExpress.XtraTreeList.Nodes.TreeListNode objNode in treeListImportSupplProducts.Nodes)
                        {
                            if (objNode.Tag == null) { continue; }
                            if (System.Convert.ToInt32(objNode.GetValue(colImportSupplQuantity)) == 0) { continue; }

                            cmd.Parameters["@PartsId"].Value = System.Convert.ToInt32(objNode.Tag);
                            cmd.Parameters["@PartsQty"].Value = System.Convert.ToInt32(objNode.GetValue(colImportSupplQuantity));
                            cmd.ExecuteNonQuery();
                            iRes = (System.Int32)cmd.Parameters["@RETURN_VALUE"].Value;
                            if (iRes == 0)
                            {
                                newRowItem = m_dsOrders.Tables["tblSupplItems"].NewRow();
                                newRowItem["SupplItem_Guid"] = (System.Guid)cmd.Parameters["@SupplItms_Guid"].Value;
                                newRowItem["Suppl_Guid"] = SupplId;
                                newRowItem["Parts_Guid"] = (System.Guid)cmd.Parameters["@Parts_Guid"].Value;
                                newRowItem["Measure_Guid"] = (System.Guid)cmd.Parameters["@Measure_Guid"].Value;
                                newRowItem["SupplItem_OrderQuantity"] = System.Convert.ToInt32(cmd.Parameters["@SplItms_OrderQty"].Value);
                                newRowItem["SupplItem_Quantity"] = System.Convert.ToInt32(cmd.Parameters["@SplItms_Quatity"].Value);

                                m_dsOrders.Tables["tblSupplItems"].Rows.Add(newRowItem);
                            }
                            else
                            {
                                SendMessageToLog((System.String)objNode.GetValue(colPartaName) + " " + (System.String)cmd.Parameters["@ERROR_MES"].Value);
                            }
                        }

                        m_dsOrders.AcceptChanges();

                        //теперь попробуем воспользоваться web сервисом
                        cmd.Parameters.Clear();
                        cmd.CommandText = System.String.Format("[{0}].[dbo].[sp_GetWebServiceName]", m_objProfile.GetOptionsDllDBName());
                        cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                        cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@SoapUrl", System.Data.SqlDbType.NVarChar, 400));
                        cmd.Parameters["@SoapUrl"].Direction = System.Data.ParameterDirection.Output;
                        cmd.ExecuteNonQuery();
                        iRes = (System.Int32)cmd.Parameters["@RETURN_VALUE"].Value;
                        System.Int32 iRetCode = -1;

                        if (iRes == 0)
                        {
                            SendMessageToLog("web сервис: " + (System.String)cmd.Parameters["@SoapUrl"].Value);
                            SendMessageToLog("идет вызов web сервиса");
                            ERPMercuryImportSuppl.WebReference.SOAPService objSrvGprs = null;
                            try
                            {
                                SendMessageToLog("v-iis01.SOAPService objSrvGprs = new SalesManager.v-iis01.SOAPService()");
                                objSrvGprs = new ERPMercuryImportSuppl.WebReference.SOAPService(); // .srvGPRS();

                                SendMessageToLog("объект \"сервис\" создан");

                                objSrvGprs.Url = (System.String)cmd.Parameters["@SoapUrl"].Value;

                                SendMessageToLog("сервису присвоен адрес: " + objSrvGprs.Url);

                                SendMessageToLog("вызывается метод objSrvGprs.SaveNewSuppl");

                                SendMessageToLog("количество таблиц в наборе" + m_dsOrders.Tables.Count.ToString());
                                SendMessageToLog("количество записей в таблице tblSuppl: " + m_dsOrders.Tables["tblSuppl"].Rows.Count.ToString());
                                SendMessageToLog("количество записей в таблице tblSupplItems: " + m_dsOrders.Tables["tblSupplItems"].Rows.Count.ToString());

                                iRetCode = objSrvGprs.SaveNewSuppl(m_dsOrders);

                                SendMessageToLog("вызов метода objSrvGprs.SaveNewSuppl завершен");
                            }
                            catch (System.Exception ferr)
                            {
                                DevExpress.XtraEditors.XtraMessageBox.Show(this, "вызов сервиса. Текст ошибки: " + ferr.Message, "Информация",
                                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);

                                SendMessageToLog("вызов сервиса. Текст ошибки: " + ferr.Message);
                            }

                            //objSrvGprs.Url = "http://192.168.7.27/MercuryPDA/SOAPService.asmx";

                            if (iRetCode == 0)
                            {
                                SendMessageToLog("заказ успешно обработан.");
                                DevExpress.XtraEditors.XtraMessageBox.Show(this, "Заказ успешно обработан.", "Информация",
                                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                            }
                            else
                            {
                                SendMessageToLog("ошибка передачи информации в БД. Код ошибки: " + iRetCode.ToString());
                                DevExpress.XtraEditors.XtraMessageBox.Show(this, "Ошибка передачи информации в БД.", "Ошибка",
                                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            }
                            objSrvGprs = null;
                        }
                        else
                        {
                            SendMessageToLog("ошибка передачи информации в БД.");
                            DevExpress.XtraEditors.XtraMessageBox.Show(this, "Ошибка передачи информации в БД. Не удалось получить адрес web сервиса.", "Ошибка",
                                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            pictureEdit1.Image = ERPMercuryImportSuppl.Properties.Resources.warning;
                        }
                    }
                }
                else
                {
                    SendMessageToLog((System.String)cmd.Parameters["@ERROR_MES"].Value);
                }

                cmd = null;
                DBConnection.Close();
            }
            catch (System.Exception f)
            {
                SendMessageToLog("SendSupplToDB. Текст ошибки: " + f.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
            return;
        }

        private void SaveSupplInDB()
        {
            try
            {
                System.Int32 CustomerId = ((cboxImportSupplCustomer.SelectedItem != null) ? ((ERP_Mercury.Common.CCustomer)cboxImportSupplCustomer.SelectedItem).InterBaseID : 0);
                System.String DepartCode = ((cboxImportSupplDepart.SelectedItem != null) ? ((ERP_Mercury.Common.CDepart)cboxImportSupplDepart.SelectedItem).DepartCode : "");
                System.String ChildDepartCode = ((cboxImportSupplChildDepart.SelectedItem != null) ? ((ERP_Mercury.Common.CChildDepart)cboxImportSupplChildDepart.SelectedItem).Code : "");
                System.String RttCode = ( ( cboxImportSupplRtt.SelectedItem != null ) ? ( (ERP_Mercury.Common.CRtt)cboxImportSupplRtt.SelectedItem ).Code : "" );
                System.String Suppl_Note = txtImportSupplDescription.Text; 
                System.Boolean Suppl_Bonus = checkImportSupplIsBonus.Checked;
                System.DateTime Suppl_DeliveryDate = dtImportSupplDeliveryDate.DateTime;

                if (CustomerId == 0)
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show(
                        "Необходимо указать клиента.", "Внимание!",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                if (DepartCode == "")
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show(
                        "Необходимо указать подразделение.", "Внимание!",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                if (RttCode == "")
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show(
                        "Необходимо указать розничную точку.", "Внимание!",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                if (treeListImportSupplProducts.AllNodesCount == 0)
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show(
                        "Необходимо указать список товара.", "Внимание!",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                if ( System.Convert.ToDecimal( treeListImportSupplProducts.GetSummaryValue(colImportSupplQuantity ) ) == 0)
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show(
                        "Количество товара в заказе должно быть больше нуля.", "Внимание!",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                SendSupplToDB(CustomerId, DepartCode, ChildDepartCode, RttCode,  Suppl_Note, Suppl_Bonus, Suppl_DeliveryDate);
            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(
                    "SaveSupplInDB.\nТекст ошибки: " + f.Message, "Ошибка",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void btnImportSupplSave_Click(object sender, EventArgs e)
        {
            SaveSupplInDB();
        }

        #endregion

        #region Отрисовка ячеек в списке
        private void treeListImportSupplProducts_CustomDrawNodeImages(object sender, DevExpress.XtraTreeList.CustomDrawNodeImagesEventArgs e)
        {
            try
            {
                if (treeListImportSupplProducts.Nodes.Count == 0) { return; }
                if (e.Node == null) { return; }
                int Y = e.SelectRect.Top + (e.SelectRect.Height - imglNodes.Images[0].Height) / 2;
                if (e.Node.Tag != null)
                {
                    try
                    {
                        if (System.Convert.ToBoolean(e.Node.GetValue(colImportSupplIsNotSingleValued)) == true)
                        {
                            e.Graphics.DrawImage(imglNodes.Images[2], new Point(e.SelectRect.X, Y));
                        }
                        else
                        {
                            e.Graphics.DrawImage(imglNodes.Images[0], new Point(e.SelectRect.X, Y));
                        }
                        e.Handled = true;
                    }
                    catch { }
                }
                else
                {
                    try
                    {
                        e.Graphics.DrawImage(imglNodes.Images[1], new Point(e.SelectRect.X, Y));
                        e.Handled = true;
                    }
                    catch { }
                }
            }
            catch (System.Exception f)
            {
                System.Windows.Forms.MessageBox.Show(null, "Ошибка отрисовки картинок в узлах\n" + f.Message, "Ошибка",
                   System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            return;
        }

        private void treeListImportSupplProducts_NodeCellStyle(object sender, DevExpress.XtraTreeList.GetCustomNodeCellStyleEventArgs e)
        {
            try
            {
                if (e.Node == treeListImportSupplProducts.FocusedNode) return;
                if (System.Convert.ToBoolean(e.Node.GetValue(colImportSupplIsNotSingleValued)) == true)
                {
                    e.Appearance.BackColor = Color.FromArgb(255, 128, 128);
                    e.Appearance.ForeColor = Color.White;
                }
            }
            catch { }
            return;
            //switch (e.Node[ [2].ToString())
            //{
            //    case "0":
            //        e.Appearance.BackColor = Color.MediumSpringGreen;
            //        break;
            //    case "1":
            //        e.Appearance.BackColor = Color.LightSkyBlue;
            //        break;
            //    case "2":
            //        e.Appearance.BackColor = Color.FromArgb(255, 128, 128);
            //        e.Appearance.ForeColor = Color.White;
            //        break;
            //}
        }

        #endregion

        #region Поиск товара по штрих-коду
        /// <summary>
        /// Возвращает список товаров для указанного штрих-кода
        /// </summary>
        /// <param name="objProfile">профайл</param>
        /// <param name="strBarcode">штрих-код</param>
        /// <param name="strErr">сообщение об ошибке</param>
        /// <returns>список товаров</returns>
        public static List<ERP_Mercury.Common.CProduct> GetProductListByBarcode(UniXP.Common.CProfile objProfile,
            System.String strBarcode, ref System.String strErr)
        {
            List<ERP_Mercury.Common.CProduct> objList = new List<ERP_Mercury.Common.CProduct>();
            System.Data.SqlClient.SqlConnection DBConnection = null;
            System.Data.SqlClient.SqlCommand cmd = null;

            DBConnection = objProfile.GetDBSource();
            if (DBConnection == null)
            {
                strErr += ("Не удалось получить соединение с базой данных.");
                return objList;
            }
            cmd = new System.Data.SqlClient.SqlCommand() { Connection = DBConnection, CommandType = System.Data.CommandType.StoredProcedure };

            try
            {
                cmd.CommandText = System.String.Format("[{0}].[dbo].[usp_GetPartsListIdByBarcode]", objProfile.GetOptionsDllDBName());
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Barcode", System.Data.SqlDbType.NVarChar, 13));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ERROR_NUM", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ERROR_MES", System.Data.SqlDbType.NVarChar, 4000));
                cmd.Parameters["@ERROR_MES"].Direction = System.Data.ParameterDirection.Output;
                cmd.Parameters["@Barcode"].Value = strBarcode;
                System.Data.SqlClient.SqlDataReader rs = cmd.ExecuteReader();
                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        objList.Add(
                            new ERP_Mercury.Common.CProduct() 
                            { 
                                ID_Ib = System.Convert.ToInt32(rs["Parts_Id"]),
                                Name = (String.Format("{0}  {1}", System.Convert.ToString(rs["Parts_Name"]), System.Convert.ToString(rs["Parts_Article"])))
                            }
                            );
                    }
                }

                rs.Close();
                rs.Dispose();
                cmd.Dispose();
                DBConnection.Close();
            }
            catch (System.Exception f)
            {
                strErr += ("Не удалось получить список товаров по штрих-коду. Текст ошибки: " + f.Message);
            }
			finally // очищаем занимаемые ресурсы
            {
            }
            return objList;
        }

        
        /// <summary>
        /// Производит поиск идентификатора товара по его штрих-коду
        /// </summary>
        /// <param name="objProfile">профайл</param>
        /// <param name="cmdSQL">SQL-команда</param>
        /// <param name="strBarcode">штрих-код</param>
        /// <param name="strErr">сообщение об ошибке</param>
        /// <param name="iErr">номер ошибки</param>
        /// <param name="iPartsId">УИ товара (InterBase)</param>
        /// <param name="strProductName">наименование товара</param>
        /// <param name="strProductArticle">артикул товара</param>
        /// <returns>код возврата хранимой процедуры</returns>
        public System.Int32 GetProductIdByBarcode(UniXP.Common.CProfile objProfile, System.Data.SqlClient.SqlCommand cmdSQL,
            System.String strBarcode, ref System.String strErr, ref System.Int32 iErr, 
            ref System.Int32 iPartsId, ref System.String strProductName, ref System.String strProductArticle )
        {
            System.Int32 iRet = -1;
            System.Data.SqlClient.SqlConnection DBConnection = null;
            System.Data.SqlClient.SqlCommand cmd = null;
            try
            {
                if (cmdSQL == null)
                {
                    DBConnection = objProfile.GetDBSource();
                    if (DBConnection == null)
                    {
                        strErr = "Не удалось получить соединение с базой данных.";
                        return iRet;
                    }
                    cmd = new System.Data.SqlClient.SqlCommand();
                    cmd.Connection = DBConnection;
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                }
                else
                {
                    cmd = cmdSQL;
                    cmd.Parameters.Clear();
                }

                cmd.CommandText = System.String.Format("[{0}].[dbo].[usp_GetPartsIdByBarcode]", objProfile.GetOptionsDllDBName());
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Barcode", System.Data.SqlDbType.NVarChar, 13));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ERROR_NUM", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Output, false, ((System.Byte)(0)), ((System.Byte)(0)), "", System.Data.DataRowVersion.Current, null));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@ERROR_MES", System.Data.SqlDbType.NVarChar, 4000));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Parts_Id", System.Data.SqlDbType.Int));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Parts_Name", System.Data.SqlDbType.NVarChar, 128));
                cmd.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Parts_Article", System.Data.SqlDbType.NVarChar, 56));
                cmd.Parameters["@ERROR_MES"].Direction = System.Data.ParameterDirection.Output;
                cmd.Parameters["@Parts_Id"].Direction = System.Data.ParameterDirection.Output;
                cmd.Parameters["@Parts_Name"].Direction = System.Data.ParameterDirection.Output;
                cmd.Parameters["@Parts_Article"].Direction = System.Data.ParameterDirection.Output;

                cmd.Parameters["@Barcode"].Value = strBarcode;
                cmd.ExecuteNonQuery();

                iRet = System.Convert.ToInt32(cmd.Parameters["@ERROR_NUM"].Value);
                iErr = System.Convert.ToInt32(cmd.Parameters["@ERROR_NUM"].Value);
                strErr = System.Convert.ToString(cmd.Parameters["@ERROR_MES"].Value);
                if( cmd.Parameters["@Parts_Id"].Value != System.DBNull.Value )
                {
                    iPartsId = System.Convert.ToInt32(cmd.Parameters["@Parts_Id"].Value);
                }
                if (cmd.Parameters["@Parts_Name"].Value != System.DBNull.Value)
                {
                    strProductName = System.Convert.ToString(cmd.Parameters["@Parts_Name"].Value);
                }
                if (cmd.Parameters["@Parts_Article"].Value != System.DBNull.Value)
                {
                    strProductArticle = System.Convert.ToString(cmd.Parameters["@Parts_Article"].Value);
                }

                if (cmdSQL == null)
                {
                    cmd.Dispose();
                    DBConnection.Close();
                }
            }
            catch (System.Exception f)
            {
                strErr = "Не удалось получить код товара по штрих-коду. Текст ошибки: " + f.Message;
            }
            return iRet;
        }


        #endregion

        #region Загрузка данных из файла
        /// <summary>
        /// Импорт приложения к заказу из файла MS Excel
        /// </summary>
        private void ImportSupplSelectFileMSExcel()
        {
            try
            {
                if (DevExpress.XtraEditors.XtraMessageBox.Show(
                    "Файл MS Excel должен содержать следующие столбцы:\n\nA - штрих-код товара\nB - количество товара\n\nПодтвердите начало операции.", "Внимание!",
                    System.Windows.Forms.MessageBoxButtons.YesNoCancel, System.Windows.Forms.MessageBoxIcon.Information) != System.Windows.Forms.DialogResult.Yes)
                {
                    return;
                }

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    this.Refresh();
                    if ((openFileDialog.FileName != "") && (System.IO.File.Exists(openFileDialog.FileName) == true))
                    {
                        txtImportSupplFileMSExcel.Text = openFileDialog.FileName;

                        ReadDataFromXLSFileForImportInLot(txtImportSupplFileMSExcel.Text);
                    }
                }
            }//try
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(
                    "btnFileOpenDialog_Click.\nТекст ошибки: " + f.Message, "Ошибка",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                ValidatePropertiesBeforeImportInDB();
            }

            return;
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ImportSupplSelectFileMSExcel();
        }

        private void btnImportSupplSelectFileMSExcel_Click(object sender, EventArgs e)
        {
            ImportSupplSelectFileMSExcel();
        }

        /// <summary>
        /// Считывает информацию из фала MS Excel
        /// </summary>
        /// <param name="strFileName">имя файла MS Excel</param>
        private void ReadDataFromXLSFileForImportInLot(System.String strFileName)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                System.IO.FileInfo newFile = new System.IO.FileInfo(strFileName);
                if (newFile.Exists == false)
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show(
                        "Ошибка экспорта в MS Excel.\n\nНе найден файл: " + strFileName, "Ошибка",
                       System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }

                treeListImportSupplProducts.CellValueChanged -= new DevExpress.XtraTreeList.CellValueChangedEventHandler(treeListImportSupplProducts_CellValueChanged);
                treeListImportSupplProducts.Nodes.Clear();
                listEditLog.Items.Clear();

                List<System.String> objBarcodeList = null;
                List<ERP_Mercury.Common.CProduct> objProductList = null;
                List<ERP_Mercury.Common.CProduct> objProductListForCell = new List<ERP_Mercury.Common.CProduct>();

                System.String strBarcode = "";
                System.String strQUANTITY = "";
                System.Decimal dblQUANTITY = 0;
                System.String strErr = "";

                System.Int32 iCurrentRow = iStartRowForImport;
                System.Int32 i = 1;


                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                    if (worksheet != null)
                    {

                        System.Boolean bStopRead = false;
                        System.Boolean bErrExists = false;
                        System.String strFrstColumn = "";

                        while (bStopRead == false)
                        {
                            bErrExists = false;

                            // пробежим по строкам и считаем информацию
                            strFrstColumn = System.Convert.ToString(worksheet.Cells[iCurrentRow, 1].Value);
                            if (strFrstColumn == "")
                            {
                                bStopRead = true;
                            }
                            else
                            {
                                objProductListForCell.Clear();
                                objBarcodeList = null;
                                objProductList = null;
                                strBarcode = System.Convert.ToString(worksheet.Cells[iCurrentRow, iColumnBarcode].Value);
                                strQUANTITY = System.Convert.ToString(worksheet.Cells[iCurrentRow, iColumnQuantity].Value);

                                dblQUANTITY = 0;

                                // код товара
                                try
                                {
                                    strErr = "";
                                    // поиск в строке штрих-кодов по шаблону
                                    objBarcodeList = DecodeBarcodeList(strBarcode);
                                    if ((objBarcodeList == null) || (objBarcodeList.Count == 0))
                                    {
                                        bErrExists = true;
                                        listEditLog.Items.Add(String.Format("{0} не удалось определить штрих-код в строке: {1} ", i, strBarcode));
                                        listEditLog.Refresh();
                                    }
                                    else
                                    {
                                        // для каждого выделенного из строки штрих-кода запрашивается список товаров
                                        foreach (System.String strBarcodeListItem in objBarcodeList)
                                        {
                                            objProductList = GetProductListByBarcode(m_objProfile, strBarcodeListItem, ref strErr);
                                            if ((objProductList != null) && (objProductList.Count > 0))
                                            {
                                                foreach (ERP_Mercury.Common.CProduct objItem in objProductList)
                                                {
                                                    objProductListForCell.Add(objItem);
                                                }
                                            }
                                        }

                                        if (objProductListForCell.Count > 0)
                                        {
                                            objProductListForCell = objProductListForCell.Distinct<ERP_Mercury.Common.CProduct>(new CProductComparer()).ToList<ERP_Mercury.Common.CProduct>();

                                            //List<Car> distinct =
                                            //  cars
                                            //  .GroupBy(car => car.CarCode)
                                            //  .Select(g => g.First())
                                            //  .ToList();
                                        }
                                    }

                                }
                                catch
                                {
                                    bErrExists = true;
                                    listEditLog.Items.Add(String.Format("{0} ошибка поиска товара по штрих-коду, штрих-код: {1}", i, strBarcode));
                                    listEditLog.Refresh();
                                }

                                if (objProductListForCell.Count == 0)
                                {
                                    bErrExists = true;
                                    listEditLog.Items.Add(String.Format("{0} товары с указанным штрих-кодом не найдены, штрих-код: {1}", System.Convert.ToString(i), strBarcode));
                                    listEditLog.Refresh();
                                }

                                // количество
                                try
                                {
                                    dblQUANTITY = System.Math.Abs( System.Convert.ToDecimal(strQUANTITY) );                                    
                                }
                                catch
                                {
                                    bErrExists = true;
                                    listEditLog.Items.Add(String.Format("{0} ошибка преобразования количества товара в числовой формат.", i));
                                    listEditLog.Refresh();
                                }
                            }

                            if ((bErrExists == false) && (bStopRead == false) && (objProductListForCell.Count > 0))
                            {
                                foreach (ERP_Mercury.Common.CProduct objProduct in objProductListForCell)
                                {
                                    treeListImportSupplProducts.AppendNode(new object[] { strBarcode, objProduct.ProductFullName, 
                                    dblQUANTITY, ( objProductListForCell.Count > 1 ),
                                    ( ( objProductListForCell.Count > 1 ) ? "штрих-коду соответствует несколько товаров" : "" ) }, null).Tag = objProduct.ID_Ib;
                                }
                                treeListImportSupplProducts.Refresh();

                                listEditLog.Items.Add(String.Format("{0} OK ", i));
                                listEditLog.Refresh();
                            }
                            else if ((bErrExists == true) && (bStopRead == false) && (objProductListForCell.Count == 0))
                            {
                                treeListImportSupplProducts.AppendNode(new object[] { strBarcode, "Товар НЕ НАЙДЕН", 
                                    dblQUANTITY, true,
                                    "Штрих-код не найден или товар помечен, как неактивный." }, null).Tag = null;

                                treeListImportSupplProducts.Refresh();
                            }

                            iCurrentRow++;
                            i++;
                            strFrstColumn = System.Convert.ToString(worksheet.Cells[iCurrentRow, 1].Value);
                            listEditLog.Refresh();

                        } //while (bStopRead == false)
                    }
                    worksheet = null;
                }


            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(
                    "Ошибка импорта данных из файла MS Excel.\n\nТекст ошибки: " + f.Message, "Ошибка",
                   System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                treeListImportSupplProducts.CellValueChanged += new DevExpress.XtraTreeList.CellValueChangedEventHandler(treeListImportSupplProducts_CellValueChanged);

                this.Cursor = System.Windows.Forms.Cursors.Default;
            }
        }
        /// <summary>
        /// Выделяет из строки список штрих-кодов
        /// </summary>
        /// <param name="strBarcodeList">входная строка</param>
        /// <returns>список штрих-кодов</returns>
        private List<System.String> DecodeBarcodeList(System.String strBarcodeList)
        {
            List<System.String> objList = new List<string>();
            try
            {
                if (strBarcodeList.Length == 0) { return objList; }
                // формируем регулярное выражение
                System.String strPattern = @"\w[0-9]{4,13}";
                System.Text.RegularExpressions.Regex rx = new Regex(strPattern, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                if (rx == null) { return objList; }

                // Define some test strings.
                System.Text.RegularExpressions.MatchCollection matches = rx.Matches(strBarcodeList);
                if (matches.Count == 0) { return objList; }

                // Report on each match.
                foreach (Match match in matches)
                {
                    if (match.Success)
                    {
                        //string strCondition = match.Groups["condition"].Value;
                        //string strValue = match.Groups["value"].Value;
                        objList.Add(match.Value);
                    }
                }
                matches = null;
                rx = null;
            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(
                "Ошибка проверки синтаксиса условия.\n\nТекст ошибки: " + f.Message, "Внимание",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            return objList;
        }

        private void treeListImportSupplProducts_CellValueChanged(object sender, DevExpress.XtraTreeList.CellValueChangedEventArgs e)
        {
            try
            {
                if ((e.Node == null) || (e.Value == null) || (e.Column != colImportSupplBarCode)) { return; }

                System.String strBarcode = System.Convert.ToString(e.Node.GetValue(colImportSupplBarCode));
                System.String strErr = "";
                System.Int32 iErr = 0;
                System.Int32 iPARTS_ID = 0;
                System.String strProductName = "";
                System.String strProductArticle = "";

                if (strBarcode.Trim().Length > 0)
                {
                    GetProductIdByBarcode(m_objProfile, null, strBarcode, ref strErr, ref iErr, ref iPARTS_ID, ref strProductName, ref strProductArticle);
                }

                if (iPARTS_ID != 0)
                {
                    e.Node.SetValue(colImportSupplProduct, (String.Format("{0} {1}", strProductName, strProductArticle)));
                    e.Node.Tag = iPARTS_ID;
                }
                else
                {
                    e.Node.SetValue(colImportSupplProduct, null);
                    e.Node.Tag = null;
                }
                
            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(
                    "treeListImportSupplProducts_CellValueChanged.\n\nТекст ошибки: " + f.Message, "Ошибка",
                   System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            return;
        }


        #endregion

        #region Экспорт в MS Excel приложения к заказу
        /// <summary>
        /// Экспорт приложения к заказу в MS Excel
        /// </summary>
        private void ExportSupplProductListToExcel()
        {
            if (treeListImportSupplProducts.Nodes.Count == 0) { return; }
            try
            {
                System.String strFileName = (Path.GetTempPath() + "\\" + System.Guid.NewGuid().ToString() + ".xlsx");

                FileInfo newFile = new FileInfo(strFileName);
                if (newFile.Exists)
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(strFileName);
                }
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    // add a new worksheet to the empty workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Приложение к заказу");
                    //Add the headers
                    worksheet.Cells[1, 1].Value = "Штрих-код";
                    worksheet.Cells[1, 2].Value = "Товар";
                    worksheet.Cells[1, 3].Value = "Количество";
                    worksheet.Cells[1, 4].Value = "Примечание";

                    using (var range = worksheet.Cells[1, 1, 1, 4])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Font.Size = 12;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                        range.Style.Font.Color.SetColor(Color.White);
                    }

                    System.Int32 iRecordNum = 0;
                    System.Int32 iCurrentRow = 2;

                    foreach (DevExpress.XtraTreeList.Nodes.TreeListNode objNode in treeListImportSupplProducts.Nodes)
                    {
                        iRecordNum++;

                        worksheet.Cells[iCurrentRow, 1].Value = System.Convert.ToString(objNode.GetValue(colImportSupplBarCode));
                        worksheet.Cells[iCurrentRow, 2].Value = System.Convert.ToString(objNode.GetValue(colImportSupplProduct));
                        worksheet.Cells[iCurrentRow, 3].Value = System.Convert.ToInt32(objNode.GetValue(colImportSupplQuantity));
                        worksheet.Cells[iCurrentRow, 4].Value = System.Convert.ToString(objNode.GetValue(colImportSupplDescription));

                        iCurrentRow++;
                    }

                    worksheet.Cells[iCurrentRow, 1].Value = "Итого:";
                    worksheet.Cells[iCurrentRow, 3, iCurrentRow, 3].FormulaR1C1 = "=SUM(R[-" + iRecordNum.ToString() + "]C:R[-1]C)";
                    worksheet.Cells[iCurrentRow, 3, iCurrentRow, 3].Style.Font.Bold = true;

                    worksheet.Cells["A1:C1000"].AutoFitColumns();

                    worksheet = null;

                    package.Save();

                    try
                    {
                        using (System.Diagnostics.Process process = new System.Diagnostics.Process())
                        {
                            process.StartInfo.FileName = strFileName;
                            process.StartInfo.Verb = "Open";
                            process.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
                            process.Start();
                        }
                    }
                    catch
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show(this, "Не удалось найти приложение, чтобы открыть файл \"" + strFileName + "\"", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(
                    "ExportSupplProductListToExcel.\nТекст ошибки: " + f.Message, "Ошибка",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            return ;
        }

        private void contextMenuStriptreeListImportSupplProducts_Opening(object sender, CancelEventArgs e)
        {
            mitemExportToExcel.Enabled = (treeListImportSupplProducts.Nodes.Count > 0);
            mitemDeleteNode.Enabled = ((treeListImportSupplProducts.Nodes.Count > 0) && (treeListImportSupplProducts.FocusedNode != null));
        }

        private void mitemExportToExcel_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                ExportSupplProductListToExcel();
            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(
                    "mitemExportToExcel_Click.\nТекст ошибки: " + f.Message, "Ошибка",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
        #endregion

        #region Запись содержимого журнала событий в файл
        /// <summary>
        /// Экспорт содержимого журнала событий в файл
        /// </summary>
        private void SaveLogInFileTXT()
        {
            if (listEditLog.Items.Count == 0) { return; }
            try
            {
                System.String strFileName = "";

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    //Stream myStream;

                    saveFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                    saveFileDialog.FilterIndex = 2;
                    saveFileDialog.RestoreDirectory = true;
                    saveFileDialog.DefaultExt = "txt";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        strFileName = saveFileDialog.FileName; 
                    }
                }

                if (strFileName != "")
                {
                    //создать (если нет) либо открыть если есть и записать текст (путем замены если что то      было   записано)
                    System.String strLog = "";
                    for( System.Int32 i = 0; i < listEditLog.Items.Count; i++ )
                    {
                        strLog += (listEditLog.GetItemText(i) + Environment.NewLine );
                    }
                    System.IO.File.WriteAllText(strFileName, strLog);
                }

            }
            catch (System.Exception f)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(
                    "SaveLogInFileTXT.\nТекст ошибки: " + f.Message, "Ошибка",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            return;
        }
        private void mitemSaveToFileTXT_Click(object sender, EventArgs e)
        {
            SaveLogInFileTXT();
        }
        private void contextMenuStripExportLogInTXT_Opening(object sender, CancelEventArgs e)
        {
            mitemSaveToFileTXT.Enabled = (listEditLog.Items.Count > 0);
        }
        #endregion

        #region Проверка указанных значений
        /// <summary>
        /// Проверка заполнения обязательных значений
        /// </summary>
        private void ValidatePropertiesBeforeImportInDB()
        {
            try
            {
                cboxImportSupplDepart.Properties.Appearance.BackColor = ((cboxImportSupplDepart.SelectedItem == null) ? System.Drawing.Color.Tomato : System.Drawing.Color.White);
                cboxImportSupplCustomer.Properties.Appearance.BackColor = ((cboxImportSupplCustomer.SelectedItem == null) ? System.Drawing.Color.Tomato : System.Drawing.Color.White);
                cboxImportSupplRtt.Properties.Appearance.BackColor = ((cboxImportSupplRtt.SelectedItem == null) ? System.Drawing.Color.Tomato : System.Drawing.Color.White);

                btnImportSupplSave.Enabled = ((cboxImportSupplDepart.SelectedItem != null) &&
                    (cboxImportSupplCustomer.SelectedItem != null) && (cboxImportSupplRtt.SelectedItem != null) && 
                    (treeListImportSupplProducts.AllNodesCount > 0));
            }
            catch (System.Exception f)
            {
                SendMessageToLog("ValidatePropertiesBeforeImportInDB. Текст ошибки: " + f.Message);
            }
            finally
            {
            }
            return;
        }

        #endregion

        #region Выход
        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
        #endregion

        #region Удаление позиции из списка
        /// <summary>
        /// Удаление позиции из приложения к заказу
        /// </summary>
        /// <param name="objNode">позиция в заказе</param>
        private void DeleteNodeFromtreeListImportSupplProducts(DevExpress.XtraTreeList.Nodes.TreeListNode objNode)
        {
            if (objNode == null) { return; }


            try
            {

                this.tableLayoutPanel3.SuspendLayout();
                ((System.ComponentModel.ISupportInitialize)(this.treeListImportSupplProducts)).BeginInit();

                System.String strBarcodeInCell = System.Convert.ToString(objNode.GetValue(colImportSupplBarCode));

                treeListImportSupplProducts.Nodes.Remove(objNode);

                if( ( strBarcodeInCell.Trim().Length > 0 ) && (treeListImportSupplProducts.Nodes.Count > 0))
                {
                    System.Int32 iNodesCount = 0;
                    foreach( DevExpress.XtraTreeList.Nodes.TreeListNode objItem in treeListImportSupplProducts.Nodes )
                    {
                        if( System.Convert.ToString( objItem.GetValue( colImportSupplBarCode ) ) == strBarcodeInCell )
                        {
                            iNodesCount++;
                        }
                    }
                    if( iNodesCount == 1 )
                    {
                        // штрих-код из удалённой позиции найден в одном узле
                        // необходимо снять пометку о неоднозначном соответствии штрих-кода товару
                        foreach (DevExpress.XtraTreeList.Nodes.TreeListNode objItem in treeListImportSupplProducts.Nodes)
                        {
                            if (System.Convert.ToString(objItem.GetValue(colImportSupplBarCode)) == strBarcodeInCell)
                            {
                                objItem.SetValue(colImportSupplIsNotSingleValued, false);
                                objItem.SetValue(colImportSupplDescription, "");
                                break;
                            }
                        }
                    }
                }
            }
            catch (System.Exception f)
            {
                SendMessageToLog("DeleteNode. Текст ошибки: " + f.Message);
            }
            finally
            {
                this.tableLayoutPanel3.ResumeLayout(false);
                ((System.ComponentModel.ISupportInitialize)(this.treeListImportSupplProducts)).EndInit();
            }
            return;
        }
        private void mitemDeleteNode_Click(object sender, EventArgs e)
        {
            DeleteNodeFromtreeListImportSupplProducts(treeListImportSupplProducts.FocusedNode);
        }

        #endregion

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            Close();
        }


        #endregion
    }

    public class EditorSupplBlank : PlugIn.IClassTypeView
    {
        public override void Run(UniXP.Common.MENUITEM objMenuItem, System.String strCaption)
        {
            frmSupplBlank obj = new frmSupplBlank(objMenuItem.objProfile, objMenuItem, enumFormOpenMode.CreateBlank);
            obj.Text = strCaption;
            obj.MdiParent = objMenuItem.objProfile.m_objMDIManager.MdiParent;
            obj.Visible = true;
        }
    }

    public class ImportSuppl : PlugIn.IClassTypeView
    {
        public override void Run(UniXP.Common.MENUITEM objMenuItem, System.String strCaption)
        {
            frmSupplBlank obj = new frmSupplBlank(objMenuItem.objProfile, objMenuItem, enumFormOpenMode.ImportSupplFromBlank);
            obj.Text = strCaption;
            obj.MdiParent = objMenuItem.objProfile.m_objMDIManager.MdiParent;
            obj.Visible = true;
        }
    }

    public class ImportSupplByBarcodes : PlugIn.IClassTypeView
    {
        public override void Run(UniXP.Common.MENUITEM objMenuItem, System.String strCaption)
        {
            frmSupplBlank obj = new frmSupplBlank(objMenuItem.objProfile, objMenuItem, enumFormOpenMode.ImportSupplFromByBarcodes);
            obj.Text = strCaption;
            obj.MdiParent = objMenuItem.objProfile.m_objMDIManager.MdiParent;
            obj.Visible = true;
        }
    }

    class CProductComparer : IEqualityComparer<ERP_Mercury.Common.CProduct>
    {
        #region IEqualityComparer<CProduct> Members

        public bool Equals(ERP_Mercury.Common.CProduct x, ERP_Mercury.Common.CProduct y)
        {
            return x.ID_Ib.Equals(y.ID_Ib);
        }

        public int GetHashCode(ERP_Mercury.Common.CProduct obj)
        {
            return obj.ID_Ib.GetHashCode();
        }

        #endregion
    }

}
