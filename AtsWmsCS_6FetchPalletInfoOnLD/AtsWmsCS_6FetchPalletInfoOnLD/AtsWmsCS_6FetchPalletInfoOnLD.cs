using log4net;
using log4net.Config;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace AtsWmsCS_6FetchPalletInfoOnLD
{
    public partial class AtsWmsCS_6FetchPalletInfoOnLD : ServiceBase
    {
        static string className = "AtsWmsCS_6FetchPalletInfoOnLD";
        private static readonly ILog Log = LogManager.GetLogger(className);
        public AtsWmsCS_6FetchPalletInfoOnLD()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {

            try
            {
                Log.Debug("OnStart :: AtsWmsA1FetchPalletInfoOnLD in OnStart....");

                try
                {
                    XmlConfigurator.Configure();
                    try
                    {
                        AtsWmsA1FetchPalletInfoOnLDTaskThread();
                    }
                    catch (Exception ex)
                    {
                        Log.Error("OnStart :: Exception occured while AtsWmsA1FetchPalletInfoOnLoadingStationTaskThread  threads task :: " + ex.Message);
                    }
                    Log.Debug("OnStart :: AtsWmsA1FetchPalletInfoOnLDTaskThread in OnStart ends..!!");
                }
                catch (Exception ex)
                {
                    Log.Error("OnStart :: Exception occured in OnStart :: " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Log.Error("OnStart :: Exception occured in OnStart :: " + ex.Message);
            }
        }
        public async void AtsWmsA1FetchPalletInfoOnLDTaskThread()
        {
            await Task.Run(() =>
            {
                try
                {
                    AtsWmsCS_6FetchPalletInfoOnLDDetails AtsWmsCS_6FetchPalletInfoOnLDDetailsInstance = new AtsWmsCS_6FetchPalletInfoOnLDDetails();
                    AtsWmsCS_6FetchPalletInfoOnLDDetailsInstance.startOperation();


                   // AtsWmsStationWorkDoneFunctionalityDetails AtsWmsStationWorkDoneFunctionalityDetailsInstance = new AtsWmsStationWorkDoneFunctionalityDetails();
                    //AtsWmsStationWorkDoneFunctionalityDetailsInstance.startOperation();

                   // AtsWmsA1FetchEngineCodeOnLDDetails AtsWmsA1FetchEngineCodeOnLDDetailsInstance = new AtsWmsA1FetchEngineCodeOnLDDetails();
                    //AtsWmsA1FetchEngineCodeOnLDDetailsInstance.startOperation();

                    //AtsWmsCS_6GiveWorkdoneOnLDDetails AtsWmsA1GiveWorkdoneOnLDDetailsInstance = new AtsWmsCS_6GiveWorkdoneOnLDDetails();
                    //AtsWmsA1GiveWorkdoneOnLDDetailsInstance.startOperation();
                }
                catch (Exception ex)
                {
                    Log.Error("TestService :: Exception in AtsWmsA1FetchPalletInfoOnLoadingStationTaskThread :: " + ex.Message);
                }

            });
        }
        protected override void OnStop()
        {
            try
            {
                Log.Debug("OnStop :: AtsWmsA1FetchPalletInfoOnLD in OnStop ends..!!");
            }
            catch (Exception ex)
            {
                Log.Error("OnStop :: Exception occured in OnStop :: " + ex.Message);
            }
        }
    }
}
