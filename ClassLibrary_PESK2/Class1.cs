using Eplan.EplApi.ApplicationFramework;
using System.Windows.Forms;
using Eplan.EplApi.MasterData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary_PESK2
{
    public class AddInModule : IEplAddIn
    {
        /*Выполняется один раз при загрузке */
        public bool OnRegister(ref bool bLoadOnStart)
        {
            bLoadOnStart = true;
            return true;
        }

        /*Выполняется один раз при выгрузке */
        public bool OnUnregister()
        {
            return true;
        }

        /*Выполняется каждый раз при запуске */
        public bool OnInit()
        {
            return true;
        }

        /*Выполняется каждый раз при прогрузке менюшек - здесь создаётся отдельная кнопка, этот код оставляю 
         В строке 
        PLC_menu.AddMenuItem("Стандарты...", "StandardsAction", "", menuId, 1, false, false);
        StandardsAction - название действия, вызываемого при тыке на кнопку
        Следовательно, создавать действе надо с этим названием
         */
        public bool OnInitGui()
        {
            Eplan.EplApi.Gui.Menu PLC_menu = new Eplan.EplApi.Gui.Menu();

            uint menuId = PLC_menu.AddMainMenu("Доп", "Сервисные программы", "Конфигуратор", "аггггг", "", int.MaxValue);
            return true;
        }

        public bool OnExit()
        {
            return true;
        }
    }
}
