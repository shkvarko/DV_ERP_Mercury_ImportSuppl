using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ERPMercuryImportSuppl
{
    public class CERPMercuryImportSupplModuleClassInfo : UniXP.Common.CModuleClassInfo
    {
        public CERPMercuryImportSupplModuleClassInfo()
        {
            m_arClassInfo.Add(new UniXP.Common.CLASSINFO() { enClassType = UniXP.Common.EnumClassType.mcView,
                                                             strClassName = "ERPMercuryImportSuppl.EditorSupplBlank", 
                strName = "Бланк заказа", strDescription = "Формирование бланка заказа для клиента", 
                lID = 0, nImage = 1, strResourceName = "summary_add_16" }
                );

            m_arClassInfo.Add(new UniXP.Common.CLASSINFO() { enClassType = UniXP.Common.EnumClassType.mcView,
                                                             strClassName = "ERPMercuryImportSuppl.ImportSuppl", 
                strName = "Импорт заказа", strDescription = "Импорт данных для формирования заказа", 
                lID = 1, nImage = 1, strResourceName = "shopping_cart_add_16" }
                );

            m_arClassInfo.Add(new UniXP.Common.CLASSINFO() { enClassType = UniXP.Common.EnumClassType.mcView,
                                                             strClassName = "ERPMercuryImportSuppl.ImportSupplByBarcodes", 
                strName = "Импорт заказа по ш/к", strDescription = "Импорт данных для формирования заказа по штрих-кодам товара", 
                lID = 2, nImage = 1, strResourceName = "shopping_cart_add_16" }
                );
         
        }
    }

    public class CERPMercuryImportSupplModuleInfo : UniXP.Common.CClientModuleInfo
    {
        public CERPMercuryImportSupplModuleInfo()
            : base(Assembly.GetExecutingAssembly(),
                UniXP.Common.EnumDLLType.typeItem,
                new System.Guid("{38AA7415-0425-4989-8DB2-83EEF103C924}"),
                new System.Guid("{A6319AD0-08C0-49ED-B25B-659BAB622B15}"),
                ERPMercuryImportSuppl.Properties.Resources.IMAGES_IMPORTSUPPLSMALL,
                ERPMercuryImportSuppl.Properties.Resources.IMAGES_IMPORTSUPPL)
        {
        }

        /// <summary>
        /// Выполняет операции по проверке правильности установки модуля в системе.
        /// </summary>
        /// <param name="objProfile">Профиль пользователя.</param>
        public override System.Boolean Check(UniXP.Common.CProfile objProfile)
        {
            return true;
        }
        /// <summary>
        /// Выполняет операции по установке модуля в систему.
        /// </summary>
        /// <param name="objProfile">Профиль пользователя.</param>
        public override System.Boolean Install(UniXP.Common.CProfile objProfile)
        {
            return true;
        }
        /// <summary>
        /// Выполняет операции по удалению модуля из системы.
        /// </summary>
        /// <param name="objProfile">Профиль пользователя.</param>
        public override System.Boolean UnInstall(UniXP.Common.CProfile objProfile)
        {
            return true;
        }
        /// <summary>
        /// Производит действия по обновлению при установке новой версии подключаемого модуля.
        /// </summary>
        /// <param name="objProfile">Профиль пользователя.</param>
        public override System.Boolean Update(UniXP.Common.CProfile objProfile)
        {
            return true;
        }
        /// <summary>
        /// Возвращает список доступных классов в данном модуле.
        /// </summary>
        public override UniXP.Common.CModuleClassInfo GetClassInfo()
        {
            return new CERPMercuryImportSupplModuleClassInfo();
        }
    }

    public class ModuleInfo : PlugIn.IModuleInfo
    {
        public UniXP.Common.CClientModuleInfo GetModuleInfo()
        {
            return new CERPMercuryImportSupplModuleInfo();
        }
    }
}
