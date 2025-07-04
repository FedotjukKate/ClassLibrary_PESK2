using System;
using System.Windows.Forms;
using Eplan.EplApi.DataModel;
using static Eplan.EplApi.DataModel.Properties;

namespace ClassLibrary_PESK2
{
    public static class Draw
    {
        public static Eplan.EplApi.DataModel.MasterData.SymbolVariant SetSymbolVariant(Eplan.EplApi.DataModel.Project oProject, string strSymbolLibName, string strSymbolName, int nVariant)
        {
            Eplan.EplApi.DataModel.MasterData.SymbolLibrary oSymbolLibrary = new Eplan.EplApi.DataModel.MasterData.SymbolLibrary(oProject, strSymbolLibName);
            Eplan.EplApi.DataModel.MasterData.Symbol oSymbol = new Eplan.EplApi.DataModel.MasterData.Symbol(oSymbolLibrary, strSymbolName);
            Eplan.EplApi.DataModel.MasterData.SymbolVariant oSymbolVariant = new Eplan.EplApi.DataModel.MasterData.SymbolVariant();
            try
            {
                oSymbolVariant.Initialize(oSymbol, nVariant);
                return oSymbolVariant;
            }

            catch (Exception e)
            {
                MessageBox.Show("Couldn't set SymbolVariant: Message: " + e.Message + "\n Source: " + e.Source + "\n Stack: " + e.StackTrace);
                return null;
            }

        }
    }
}
