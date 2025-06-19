using Eplan.EplApi.ApplicationFramework;
using System.Windows.Forms;

namespace ClassLibrary_PESK2.StandardsAction
{
    /*
     * Вызывает окошко просмотра страндартов. Привязано к менюшке Pesk ->Стандарты
     * */
    public class StandardAction : IEplAction
    {
        /*Создание действия; ordinal и имена у каждого должны быть уникален - мною заняты 22, 25,28
         * (я в курсе, что это верх логики, ешьте меня)     
         */
        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = "аггггг";
            Ordinal = 30;
            return true;
        }

        /*Собстна действие*/
        public bool Execute(ActionCallingContext oActionCallingContext)
        {
            Form_PESK form1 = new Form_PESK();
            form1.Show();
            MessageBox.Show("Х*ета");
            return true;
        }

        /*я хз чё это, но оно должно быть*/
        public void GetActionProperties(ref ActionProperties actionProperties)
        {

        }
    }
}
