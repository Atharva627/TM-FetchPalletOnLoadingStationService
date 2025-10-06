using AtsWmsCS_6FetchPalletInfoOnLD.ats_tata_metallics_dbDataSetTableAdapters;
using log4net;
using OPCAutomation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading.Tasks;
using static AtsWmsCS_6FetchPalletInfoOnLD.ats_tata_metallics_dbDataSet;

namespace AtsWmsCS_6FetchPalletInfoOnLD
{
    class AtsWmsCS_6FetchPalletInfoOnLDDetails
    {

        #region DataTables
        ats_wms_master_plc_connection_detailsDataTable ats_wms_master_plc_connection_detailsDataTableDT = null;
        ats_wms_master_pallet_informationDataTable ats_wms_master_pallet_informationDataTableDT = null;
        ats_wms_master_pallet_informationDataTable ats_wms_master_pallet_informationDataTableDTDataInsert = null;
        ats_wms_infeed_mission_runtime_detailsDataTable ats_wms_infeed_mission_runtime_detailsDataTableDT = null;
        ats_wms_master_product_detailsDataTable ats_wms_master_product_detailsDataTableDT = null;
        ats_wms_loading_stations_tag_detailsDataTable ats_wms_loading_stations_tag_detailsDataTableDT = null;
        ats_wms_loading_stations_tag_detailsDataTable ats_wms_loading_stations_tag_detailsDataTableIsActiveDT = null;
        ats_wms_master_station_detailsDataTable ats_wms_master_station_detailsDataTableDT = null;
        ats_wms_master_equipment_detailsDataTable ats_wms_master_equipment_detailsDataTableDT = null;

        #endregion

        #region TableAdapters
        ats_wms_master_plc_connection_detailsTableAdapter ats_wms_master_plc_connection_detailsTableAdapterInstance = new ats_wms_master_plc_connection_detailsTableAdapter();
        ats_wms_master_pallet_informationTableAdapter ats_wms_master_pallet_informationTableAdapterInstance = new ats_wms_master_pallet_informationTableAdapter();
        ats_wms_infeed_mission_runtime_detailsTableAdapter ats_wms_infeed_mission_runtime_detailsTableAdapterInstance = new ats_wms_infeed_mission_runtime_detailsTableAdapter();
        ats_wms_master_product_detailsTableAdapter ats_wms_master_product_detailsTableAdapterInstance = new ats_wms_master_product_detailsTableAdapter();
        ats_wms_loading_stations_tag_detailsTableAdapter ats_wms_loading_stations_tag_detailsTableAdapterInstance = new ats_wms_loading_stations_tag_detailsTableAdapter();
        ats_wms_master_station_detailsTableAdapter ats_wms_master_station_detailsTableAdapterInstance = new ats_wms_master_station_detailsTableAdapter();
        ats_wms_master_equipment_detailsTableAdapter ats_wms_master_equipment_detailsTableAdapterInstance = new ats_wms_master_equipment_detailsTableAdapter();

        #endregion

        #region PLC PING VARIABLE   
        private Ping pingSenderForThisConnection = null;
        private PingReply replyForThisConnection = null;
        private Boolean pingStatus = false;
        private int serverPingStatusCount = 0;
        #endregion

        #region KEPWARE VARIABLES

        /* Kepware variable*/

        OPCServer ConnectedOpc = new OPCServer();

        Array OPCItemIDs = Array.CreateInstance(typeof(string), 100);
        Array ItemServerHandles = Array.CreateInstance(typeof(Int32), 100);
        Array ItemServerErrors = Array.CreateInstance(typeof(Int32), 100);
        Array ClientHandles = Array.CreateInstance(typeof(Int32), 100);
        Array RequestedDataTypes = Array.CreateInstance(typeof(Int16), 100);
        Array AccessPaths = Array.CreateInstance(typeof(string), 100);
        Array ItemServerValues = Array.CreateInstance(typeof(string), 100);
        OPCGroup OpcGroupNames;
        //object yDIR;
        // object zDIR;
        object yDIR;
        object zDIR;

        // Connection string
        static string plcServerConnectionString = null;

        #endregion

        #region Global Variables
        static string className = "AtsWmsCS_6FetchPalletInfoOnLDDetails";
        private static readonly ILog Log = LogManager.GetLogger(className);
        private System.Timers.Timer a1FetchPalletInfoOnLDTimer = null;
        private string IP_ADDRESS = "192.168.0.1";
        int stationId = 0;
        string CS_PalletCode = "";
        string coreQuantity = "";
        string coreMaterialCode = "";
        string palletPresentOnPickup = "";
        string RobotRowTag = "";
        string RobotColumnTag = "";
        string RobotLayerTag = "";
        string palletCodeOnPickup = "";
        string palletMatrialCodeOnPickup = "";
        float qtyOnPickup;
        string readyToPickup = "";
        int palletStatusId = 0;
        string palletStatusName = "";
        #endregion



        public void startOperation()
        {
            Log.Debug("1");
            try
            {
                //Timer 
                a1FetchPalletInfoOnLDTimer = new System.Timers.Timer();
                //Running the function after 1 sec 
                a1FetchPalletInfoOnLDTimer.Interval = (1000);
                //to reset timer after completion of 1 cycle

                a1FetchPalletInfoOnLDTimer.Elapsed += new System.Timers.ElapsedEventHandler(atsWmsCS_6FetchPalletInfoOnLDDetailsOperation);
                a1FetchPalletInfoOnLDTimer.AutoReset = false;
                //Enabling the timer
                //a1FetchPalletInfoOnLDTimer.Enabled = true;
                //Timer Start
                //a1FetchPalletInfoOnLDTimer.Start();
                //After 1 sec timer will elapse and DataFetchDetailsOperation function will be called 

                a1FetchPalletInfoOnLDTimer.Start();
            }
            catch (Exception ex)
            {
                Log.Error("startOperation :: Exception Occure in a1FetchPalletInfoOnLoadingStationTimer" + ex.Message);
            }
        }




        public void atsWmsCS_6FetchPalletInfoOnLDDetailsOperation(object sender, EventArgs args)
        {
            try
            {
                Log.Debug("2");
                try
                {
                    //Stopping the timer to start the below operation
                    a1FetchPalletInfoOnLDTimer.Stop();
                }
                catch (Exception ex)
                {
                    Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception occure while stopping the timer :: " + ex.Message + "StackTrace  :: " + ex.StackTrace);
                }

                try
                {
                    //Fetching PLC data from DB by sending PLC connection IP address
                    ats_wms_master_plc_connection_detailsDataTableDT = ats_wms_master_plc_connection_detailsTableAdapterInstance.GetDataByPLC_CONNECTION_IP_ADDRESS(IP_ADDRESS);
                    Log.Debug("3");
                }
                catch (Exception ex)
                {
                    Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception Occure while reading machine datasource connection details from the database :: " + ex.Message + "StackTrace :: " + ex.StackTrace);
                }


                // Check PLC Ping Status
                try
                {
                    Log.Debug("4");
                    //Checking the PLC ping status by a method
                    pingStatus = checkPlcPingRequest();
                }
                catch (Exception ex)
                {
                    Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while checking plc ping status :: " + ex.Message + " stactTrace :: " + ex.StackTrace);
                }

                if (pingStatus == true)
                //if (true)
                {
                    try
                    {
                        Log.Debug("5");
                        //checking if the PLC data from DB is retrived or not
                        if (ats_wms_master_plc_connection_detailsDataTableDT != null && ats_wms_master_plc_connection_detailsDataTableDT.Count != 0)
                        //if (true)
                        {
                            try
                            {
                                plcServerConnectionString = ats_wms_master_plc_connection_detailsDataTableDT[0].PLC_CONNECTION_URL;
                            }
                            catch (Exception ex)
                            {
                                Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while getting plcServerConnectionString :: " + ex.Message + " stackTrace :: " + ex.StackTrace);
                            }
                            try
                            {
                                //Calling the connection method for PLC connection
                                OnConnectPLC();
                            }
                            catch (Exception ex)
                            {
                                Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while connecting to plc :: " + ex.Message + " stackTrace :: " + ex.StackTrace);
                            }
                            try
                            {

                                // Check the PLC connected status
                                if (ConnectedOpc.ServerState.ToString().Equals("1"))
                                //if (true)
                                {
                                    //BussinessLogic
                                    try
                                    {
                                        //getting loading station details
                                        ats_wms_master_station_detailsDataTableDT = ats_wms_master_station_detailsTableAdapterInstance.GetDataBySTATION_TYPE("LOADING");

                                        if (ats_wms_master_station_detailsDataTableDT != null && ats_wms_master_station_detailsDataTableDT.Count > 0)
                                        {
                                            for (int i = 0; i < ats_wms_master_station_detailsDataTableDT.Count; i++)
                                            {


                                                try
                                                {

                                                }
                                                catch (Exception)
                                                {

                                                    throw;
                                                }
                                                Log.Debug("6 :: Fetching Loading Station Details");
                                                //getting station data one by one
                                                ats_wms_loading_stations_tag_detailsDataTableDT = ats_wms_loading_stations_tag_detailsTableAdapterInstance.GetDataBySTATION_ID(ats_wms_master_station_detailsDataTableDT[i].STATION_ID);

                                                Log.Debug("6.1 :: Station Id recieval Check");
                                                if (ats_wms_loading_stations_tag_detailsDataTableDT != null && ats_wms_loading_stations_tag_detailsDataTableDT.Count > 0)
                                                {
                                                    Log.Debug("6.1.1 :: ");
                                                    //Checking if Station is Active
                                                    ats_wms_loading_stations_tag_detailsDataTableIsActiveDT = ats_wms_loading_stations_tag_detailsTableAdapterInstance.GetDataBySTATION_IS_ACTIVEAndTAG_DETAILS_ID(1, ats_wms_loading_stations_tag_detailsDataTableDT[0].TAG_DETAILS_ID);

                                                    //Log.Debug("6.2 :: The Active Station name is: "+ ats_wms_loading_stations_tag_detailsDataTableIsActiveDT[0].STATION_NAME);
                                                    if (ats_wms_loading_stations_tag_detailsDataTableIsActiveDT != null && ats_wms_loading_stations_tag_detailsDataTableIsActiveDT.Count > 0)
                                                    {

                                                        Log.Debug("7 :: The Station is Active");

                                                        if (ats_wms_loading_stations_tag_detailsDataTableIsActiveDT[0].IS_ACTIVE_TAG.Equals("True"))
                                                        {


                                                            //checking if the pallet is present
                                                            palletPresentOnPickup = readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].PICKUP_POSITION_PALLET_PRESENT_TAG);

                                                            if (palletPresentOnPickup.Equals("True"))
                                                            {
                                                                Log.Debug("8 :: Pallet is present at Loading Station");
                                                                //reading pallet code
                                                                palletCodeOnPickup = readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].PICKUP_POSITION_PALLET_CODE_TAG).Trim();

                                                                if (palletCodeOnPickup != null && palletCodeOnPickup != "" && palletCodeOnPickup != "0")
                                                                {
                                                                    //format pallet code
                                                                    #region Formating pallet code
                                                                    //Checking if the pallet code length is greater than 8 (as define by costumer standerds)
                                                                    if (palletCodeOnPickup.Length > 4)
                                                                    {
                                                                        try
                                                                        {
                                                                            //Extracting and Storing last 8 digit from scanned pallet code
                                                                            palletCodeOnPickup = palletCodeOnPickup.Substring(palletCodeOnPickup.Length - 4, 4);
                                                                            Log.Debug("AtsWmsStationCurrentPalletDetailsOperation :: Modified Pallet Code :: " + palletCodeOnPickup);
                                                                        }
                                                                        catch (Exception ex)
                                                                        {
                                                                            Log.Error("AtsWmsStationCurrentPalletDetailsOperation :: Exception occured while substring scanned pallet code ::" + ex.Message + " StackTrace:: " + ex.StackTrace);
                                                                        }
                                                                    }
                                                                    #endregion
                                                                    //checking RobotRowTag
                                                                    Log.Debug("9 :: Reading Pallet Code and Checking its Format");

                                                                    RobotRowTag = readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].ROBOT_ROW_TAG);

                                                                    if (RobotRowTag.Equals("True"))
                                                                    {
                                                                        Log.Debug("9.1 :: Reading ROBOT_ROW_TAG");

                                                                        RobotColumnTag = readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].ROBOT_COLUMN_TAG);

                                                                        if (RobotColumnTag.Equals("True"))
                                                                        {
                                                                            Log.Debug("9.2 :: Reading ROBOT_COLUMN_TAG");
                                                                            RobotLayerTag = readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].ROBOT_LAYER_TAG);

                                                                            if (RobotLayerTag.Equals("True"))
                                                                            {
                                                                                Log.Debug("9.3 :: Reading ROBOT_LAYER_TAG");

                                                                                //checking if the pallet information details are already avalilable if no inserting the new one
                                                                                ats_wms_master_pallet_informationDataTableDT = ats_wms_master_pallet_informationTableAdapterInstance.GetDataByPALLET_CODEOrderByPALLET_INFORMATION_IDDESC(palletCodeOnPickup);

                                                                                if (ats_wms_master_pallet_informationDataTableDT != null && ats_wms_master_pallet_informationDataTableDT.Count > 0)
                                                                                {
                                                                                    //update details in the station tag tables 
                                                                                    //ready material code & QTY from PLC tag
                                                                                    //check if pichup is ready
                                                                                    Log.Debug("10 :: Checking for Data entry in DB for the Pallet Code");

                                                                                    //checking if the pallet is ready for pickup
                                                                                    readyToPickup = readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].PICKUP_IS_READY_TAG);

                                                                                    if (readyToPickup.Equals("True"))
                                                                                    {
                                                                                        for (;;)
                                                                                        {
                                                                                            Log.Debug("11 :: Reading Pallet Code at Pickup Position");
                                                                                            //getting pallet material and qty details from PLC
                                                                                            palletMatrialCodeOnPickup = readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].PICKUP_POSITION_MATERIAL_CODE_TAG).Trim();

                                                                                            Log.Debug("12 :: Reading Pickup Position Quantity");

                                                                                            qtyOnPickup = float.Parse(readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].PICKUP_POSITION_QUANTITY_VALUE_TAG));

                                                                                            //checking if the values are not empty
                                                                                            if (palletMatrialCodeOnPickup != null && palletMatrialCodeOnPickup != "" && palletMatrialCodeOnPickup != "0" && qtyOnPickup > 0.0f)
                                                                                            {
                                                                                                Log.Debug("13 :: Values Entered are valid");

                                                                                                //checking material details on the pallet
                                                                                                ats_wms_master_product_detailsDataTableDT = ats_wms_master_product_detailsTableAdapterInstance.GetDataByCORE_SIZE(palletMatrialCodeOnPickup);

                                                                                                if (ats_wms_master_product_detailsDataTableDT != null && ats_wms_master_product_detailsDataTableDT.Count > 0)
                                                                                                {

                                                                                                    Log.Debug("14 :: Checked Material details of Pallet code");


                                                                                                    //updating in station tag table by checking if already updated
                                                                                                    ats_wms_loading_stations_tag_detailsDataTableDT = ats_wms_loading_stations_tag_detailsTableAdapterInstance.GetDataBySTATION_ID(ats_wms_master_station_detailsDataTableDT[i].STATION_ID);
                                                                                                    if (ats_wms_loading_stations_tag_detailsDataTableDT != null && ats_wms_loading_stations_tag_detailsDataTableDT.Count > 0)
                                                                                                    {
                                                                                                        ////Updating the Equipment_is_Active in Master Equipment Table
                                                                                                        //Log.Debug("7.1 :: Equipment that is going to be used "+ ats_wms_loading_stations_tag_detailsDataTableDT[i].STATION_NAME);

                                                                                                        //ats_wms_master_equipment_detailsTableAdapterInstance.UpdateEQUIPMENT_IS_ACTIVEwhereEQUIPMENT_NAME(1, ats_wms_loading_stations_tag_detailsDataTableDT[i].STATION_NAME);

                                                                                                        if (ats_wms_loading_stations_tag_detailsDataTableDT[0].DROP_POSITION_PALLET_CODE_VALUE != palletCodeOnPickup)
                                                                                                        {

                                                                                                            Log.Debug("15 :: Updating new Pallet information in Loading Station");

                                                                                                            //updating with new pallet details on the station table
                                                                                                            ats_wms_loading_stations_tag_detailsTableAdapterInstance.UpdateStationPalletDetails(1, palletCodeOnPickup, palletMatrialCodeOnPickup, qtyOnPickup, RobotRowTag, RobotColumnTag, RobotLayerTag, ats_wms_loading_stations_tag_detailsDataTableDT[0].TAG_DETAILS_ID);
                                                                                                            Log.Debug("15.1 :: Updated new Pallet information in Loading Station");
                                                                                                        }
                                                                                                    }

                                                                                                    Log.Debug("15.2 :: Updating new Pallet information in Loading Station");
                                                                                                    //fetching pallet info for checking if the pallet information is already updated or not
                                                                                                    ats_wms_master_pallet_informationDataTableDT = ats_wms_master_pallet_informationTableAdapterInstance.GetDataByPALLET_CODEOrderByPALLET_INFORMATION_IDDESC(palletCodeOnPickup);
                                                                                                    if (ats_wms_master_pallet_informationDataTableDT != null && ats_wms_master_pallet_informationDataTableDT.Count > 0)
                                                                                                    {

                                                                                                        Log.Debug("16 :: FEtching Pallet info for the Pallet Code ");

                                                                                                        try
                                                                                                        {
                                                                                                            if (ats_wms_master_pallet_informationDataTableDT[0].QUANTITY == 0)
                                                                                                            {

                                                                                                                Log.Debug("17 :: Updating information against the Pallet Code");

                                                                                                                //update in pallet information table
                                                                                                                ats_wms_master_pallet_informationTableAdapterInstance.UpdatePalletInfo(ats_wms_master_product_detailsDataTableDT[0].PRODUCT_ID,
                                                                                                                    ats_wms_master_product_detailsDataTableDT[0].PRODUCT_NAME,
                                                                                                                    palletMatrialCodeOnPickup, ats_wms_master_product_detailsDataTableDT[0].PRODUCT_ID, palletMatrialCodeOnPickup, palletMatrialCodeOnPickup, "NA",
                                                                                                                    ats_wms_master_station_detailsDataTableDT[i].STATION_ID,
                                                                                                                    qtyOnPickup,
                                                             
                                                                                                                    1,
                                                                                                                    "LOADED",
                                                                                                                    0,
                                                                                                                    0,
                                                                                                                    DateTime.Now.ToString(),
                                                                                                                    "NA",
                                                                                                                    1,
                                                                                                                    0, ats_wms_master_station_detailsDataTableDT[i].STATION_NAME,
                                                                                                                    ats_wms_master_pallet_informationDataTableDT[0].PALLET_INFORMATION_ID
                                                                                                                    );
                                                                                                                Log.Debug("18 :: Pallet is Loaded Successfully");

                                                                                                            }
                                                                                                            break;

                                                                                                        }
                                                                                                        catch (Exception ex)
                                                                                                        {


                                                                                                            Log.Debug(" exception occure when try to update data:: ");
                                                                                                        }
                                                                                                    }
                                                                                                    else
                                                                                                    {
                                                                                                        Log.Debug(" :: NO pallet imnformation id is found");
                                                                                                    }

                                                                                                }
                                                                                            }
                                                                                        }
                                                                                    }


                                                                                }
                                                                                else
                                                                                {
                                                                                    Log.Debug("2.1 :: Pallet Details not present in the DB.... Inserting New Information");
                                                                                    //pallet details are not into database already inserting new 
                                                                                    ats_wms_master_pallet_informationTableAdapterInstance.Insert(palletCodeOnPickup, 0, "NA", "0", 0, "NA", "NA", "NA", 0,
                                                                                       "NA", 0, 3, "Empty", 0, 0, DateTime.Now, 1, "master_user", 0, "0", 0, "NA", 0, 0, 0, 0);
                                                                                }
                                                                            }

                                                                            else
                                                                            {
                                                                                Log.Debug("ROBOT_ROW_TAG NOT FOUND");
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            Log.Debug("ROBOT_COLUMN_TAG NOT FOUND");
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        Log.Debug("15.3 :: Pallet Present is False... resettings all the Values ");
                                                                        //as pallet present is false reseting value to 0 and NA after checking if value are not already reset 
                                                                        ats_wms_loading_stations_tag_detailsDataTableDT = ats_wms_loading_stations_tag_detailsTableAdapterInstance.GetDataBySTATION_ID(ats_wms_master_station_detailsDataTableDT[i].STATION_ID);
                                                                        if (ats_wms_loading_stations_tag_detailsDataTableDT != null && ats_wms_loading_stations_tag_detailsDataTableDT.Count > 0)
                                                                        {
                                                                            if (ats_wms_loading_stations_tag_detailsDataTableDT[0].DROP_POSITION_PALLET_CODE_VALUE != "NA")
                                                                            {
                                                                                ats_wms_loading_stations_tag_detailsTableAdapterInstance.UpdateResetData(0, 0, "NA", "NA", 0, ats_wms_loading_stations_tag_detailsDataTableDT[0].TAG_DETAILS_ID);
                                                                                Log.Debug("15.3.1 :: Values Reset");
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    Log.Debug("7.2 :: The Station is Not Active in CC");
                                                                }
                                                            }
                                                            else
                                                            {
                                                                Log.Debug("7.1 :: The Station is not Active");
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    catch (Exception ex)
                                    {
                                        Log.Error("Exception occurred while reading pallet present :: " + ex.StackTrace);
                                    }
                                }
                                else
                                {
                                    //Reconnect to plc
                                }
                            }
                            catch (Exception ex)
                            {
                                Log.Error("startOperation :: Exception occured while checking server state :: " + ex.Message + " stackTrace :: " + ex.StackTrace);
                            }
                        }
                        else
                        {
                            //Reconnect to plc, Check Ip address, url
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Error("startOperation :: Exception occured while checking plc DT :: " + ex.Message + " stackTrace :: " + ex.StackTrace);
                    }
                }

            }
            catch (Exception ex)
            {
                Log.Error("startOperation :: Exception occured while stopping timer :: " + ex.Message + " stackTrace :: " + ex.StackTrace);
            }
            finally
            {

                try
                {
                    //Starting the timer again for the next iteration
                    a1FetchPalletInfoOnLDTimer.Start();
                }
                catch (Exception ex1)
                {
                    Log.Error("startOperation :: Exception occured while stopping timer :: " + ex1.Message + " stackTrace :: " + ex1.StackTrace);
                }

            }

        }


        public Tuple<string, string, string> GetPalletPresentAndPalletCode()
        {
            string palletPresent1 = "";
            string palletCode1 = "";
            string palletPresent2 = "";
            string palletCode2 = "";
            string palletPresent3 = "";
            string palletCode3 = "";
            string coreShooter = "";

            try
            {
                palletPresent1 = readTag("ATS.WMS_STACKER_1.CS_4_PICKUP_POSITION_PALLET_PRESENT");
            }
            catch (Exception ex)
            {
                Log.Error("Exception occurred reading data in ATS.WMS_STACKER_1.CS_4_PICKUP_POSITION_PALLET_PRESENT :: " + ex.Message + " StackTrace:: " + ex.StackTrace);
            }

            if (palletPresent1.Equals("True"))
            {
                try
                {
                    palletCode1 = readTag("ATS.WMS_STACKER_1.CS_4_PICKUP_POSITION_PALLET_CODE").Trim();

                    coreShooter = "CORE_SHOOTER-4";
                }
                catch (Exception ex)
                {
                    Log.Error("Exception occurred reading data in ATS.WMS_STACKER_1.AREA_1_LOADING_STATION_PALLET_CODE :: " + ex.Message + " StackTrace:: " + ex.StackTrace);
                }
                return Tuple.Create(palletPresent1, palletCode1, coreShooter);
            }

            try
            {
                palletPresent2 = readTag("ATS.WMS_STACKER_1.CS_5_PICKUP_POSITION_PALLET_PRESENT");
            }
            catch (Exception ex)
            {
                Log.Error("Exception occurred reading data in ATS.WMS_STACKER_1.CS_5_PICKUP_POSITION_PALLET_PRESENT :: " + ex.Message + " StackTrace:: " + ex.StackTrace);
            }

            if (palletPresent2.Equals("True"))
            {
                try
                {
                    palletCode2 = readTag("ATS.WMS_STACKER_1.CS_5_PICKUP_POSITION_PALLET_CODE").Trim();

                    coreShooter = "CORE_SHOOTER-5";
                }
                catch (Exception ex)
                {
                    Log.Error("Exception occurred reading data in ATS.WMS_AREA_2.AREA_2_LOADING_STATION_PALLET_CODE :: " + ex.Message + " StackTrace:: " + ex.StackTrace);
                }
                return Tuple.Create(palletPresent2, palletCode2, coreShooter);
            }

            try
            {
                palletPresent3 = readTag("ATS.WMS_STACKER_1.CS_6_PICKUP_POSITION_PALLET_PRESENT");
            }
            catch (Exception ex)
            {
                Log.Error("Exception occurred reading data in ATS.WMS_STACKER_1.CS_6_PICKUP_POSITION_PALLET_PRESENT :: " + ex.Message + " StackTrace:: " + ex.StackTrace);
            }

            if (palletPresent3.Equals("True"))
            {
                try
                {
                    palletCode3 = readTag("ATS.WMS_STACKER_1.CS_6_PICKUP_POSITION_PALLET_CODE").Trim();

                    coreShooter = "CORE_SHOOTER-6";
                }
                catch (Exception ex)
                {
                    Log.Error("Exception occurred reading data in ATS.WMS_AREA_3.AREA_3_LOADING_STATION_PALLET_CODE :: " + ex.Message + " StackTrace:: " + ex.StackTrace);
                }
                return Tuple.Create(palletPresent3, palletCode3, coreShooter);
            }

            return Tuple.Create("", "", "");
        }


        //public string checkPalletPresenceAtEmptyPlacePosition()
        //{
        //    string palletPresentCS6Empty = "";
        //    string palletPresentCS5Empty = "";
        //    string palletPresentCS4Empty = "";

        //    try
        //    {
        //        palletPresentCS6Empty = readTag("ATS.WMS_STACKER_1.CS_6_PICKUP_POSITION_PALLET_PRESENT");
        //        if (palletPresentCS6Empty.Equals("False"))
        //        {
        //            return palletPresentCS6Empty;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Log.Error("Exception occurred reading data in ATS.WMS_STACKER_1.CS_4_PICKUP_POSITION_PALLET_PRESENT :: " + ex.Message + " StackTrace:: " + ex.StackTrace);
        //    }

        //    try
        //    {
        //        palletPresentCS5Empty = readTag("ATS.WMS_STACKER_1.CS_5_PICKUP_POSITION_PALLET_PRESENT");
        //        if (palletPresentCS5Empty.Equals("False"))
        //        {
        //            return palletPresentCS5Empty;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Log.Error("Exception occurred reading data in ATS.WMS_STACKER_1.CS_4_PICKUP_POSITION_PALLET_PRESENT :: " + ex.Message + " StackTrace:: " + ex.StackTrace);
        //    }

        //    try
        //    {
        //        palletPresentCS4Empty = readTag("ATS.WMS_STACKER_1.CS_4_PICKUP_POSITION_PALLET_PRESENT");
        //        if (palletPresentCS4Empty.Equals("False"))
        //        {
        //            return palletPresentCS4Empty;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Log.Error("Exception occurred reading data in ATS.WMS_STACKER_1.CS_4_PICKUP_POSITION_PALLET_PRESENT :: " + ex.Message + " StackTrace:: " + ex.StackTrace);
        //    }

        //    return "True"; // Return "False" if no pallet is present in any area
        //}


        #region Ping funcationality

        public Boolean checkPlcPingRequest()
        {
            //Log.Debug("IprodPLCMachineXmlGenOperation :: Inside checkServerPingRequest");

            try
            {
                try
                {
                    pingSenderForThisConnection = new Ping();
                    replyForThisConnection = pingSenderForThisConnection.Send(IP_ADDRESS);
                }
                catch (Exception ex)
                {
                    Log.Error("checkPlcPingRequest :: for IP :: " + IP_ADDRESS + " Exception occured while sending ping request :: " + ex.Message + " stackTrace :: " + ex.StackTrace);
                    replyForThisConnection = null;
                }

                if (replyForThisConnection != null && replyForThisConnection.Status == IPStatus.Success)
                {
                    //Log.Debug("checkPlcPingRequest :: for IP :: " + IP_ADDRESS + " Ping success :: " + replyForThisConnection.Status.ToString());
                    return true;
                }
                else
                {
                    //Log.Debug("checkPlcPingRequest :: for IP :: " + IP_ADDRESS + " Ping failed. ");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log.Error("checkPlcPingRequest :: for IP :: " + IP_ADDRESS + " Exception while checking ping request :: " + ex.Message + " stackTrace :: " + ex.StackTrace);
                return false;
            }
        }

        #endregion

        #region Read and Write PLC tag

        [HandleProcessCorruptedStateExceptions]
        public string readTag(string tagName)
        {

            try
            {
                //Log.Debug("IprodPLCCommunicationOperation :: Inside readTag.");

                // Set PLC tag
                OPCItemIDs.SetValue(tagName, 1);
                //Log.Debug("readTag :: Plc tag is configured for plc group.");

                // remove all group
                ConnectedOpc.OPCGroups.RemoveAll();
                //Log.Debug("readTag :: Remove all group.");

                // Kepware configuration                
                OpcGroupNames = ConnectedOpc.OPCGroups.Add("AtsWmsCS_6FetchPalletInfoOnLDDetailsGroup");
                OpcGroupNames.DeadBand = 0;
                OpcGroupNames.UpdateRate = 100;
                OpcGroupNames.IsSubscribed = true;
                OpcGroupNames.IsActive = true;
                OpcGroupNames.OPCItems.AddItems(1, ref OPCItemIDs, ref ClientHandles, out ItemServerHandles, out ItemServerErrors, RequestedDataTypes, AccessPaths);
                //Log.Debug("readTag :: Kepware properties configuration is complete.");

                // Read tag
                OpcGroupNames.SyncRead((short)OPCAutomation.OPCDataSource.OPCDevice, 1, ref
                   ItemServerHandles, out ItemServerValues, out ItemServerErrors, out yDIR, out zDIR);

                //Log.Debug("readTag ::  tag name :: " + tagName + " tag value :: " + Convert.ToString(ItemServerValues.GetValue(1)));

                if (Convert.ToString(ItemServerValues.GetValue(1)).Equals("True"))
                {
                    Log.Debug("readTag :: Found and Return True");
                    Log.Debug("readTag ::  tag name :: " + tagName + " tag value :: " + Convert.ToString(ItemServerValues.GetValue(1)));
                    //ConnectedOpc.OPCGroups.Remove("AtsWmsCS_6FetchPalletInfoOnLDDetailsGroup");
                    return "True";
                }
                else if (Convert.ToString(ItemServerValues.GetValue(1)).Equals("False"))
                {
                    Log.Debug("readTag :: Found and Return False");
                    Log.Debug("readTag ::  tag name :: " + tagName + " tag value :: " + Convert.ToString(ItemServerValues.GetValue(1)));
                    //ConnectedOpc.OPCGroups.Remove("AtsWmsCS_6FetchPalletInfoOnLDDetailsGroup");
                    return "False";
                }
                else
                {
                    Log.Debug("readTag ::  tag name :: " + tagName + " tag value :: " + Convert.ToString(ItemServerValues.GetValue(1)));
                    //ConnectedOpc.OPCGroups.Remove("AtsWmsCS_6FetchPalletInfoOnLDDetailsGroup");
                    return Convert.ToString(ItemServerValues.GetValue(1));
                }

            }
            catch (Exception ex)
            {
                Log.Error("readTag :: Exception while reading plc tag :: " + tagName + " :: " + ex.Message);
            }

            Log.Debug("readTag :: Return False.. retun null.");

            return "False";
        }

        [HandleProcessCorruptedStateExceptions]
        public Boolean writeTag(string tagName, string tagValue)
        {

            try
            {
                Log.Debug("IprodGiveMissionToStacker :: Inside writeTag.");

                // Set PLC tag
                OPCItemIDs.SetValue(tagName, 1);
                //Log.Debug("writeTag :: Plc tag is configured for plc group.");

                // remove all group
                ConnectedOpc.OPCGroups.RemoveAll();
                //Log.Debug("writeTag :: Remove all group.");

                // Kepware configuration                  
                OpcGroupNames = ConnectedOpc.OPCGroups.Add("AtsWmsCS_6FetchPalletInfoOnLDDetailsGroup");
                OpcGroupNames.DeadBand = 0;
                OpcGroupNames.UpdateRate = 100;
                OpcGroupNames.IsSubscribed = true;
                OpcGroupNames.IsActive = true;
                OpcGroupNames.OPCItems.AddItems(1, ref OPCItemIDs, ref ClientHandles, out ItemServerHandles, out ItemServerErrors, RequestedDataTypes, AccessPaths);
                //Log.Debug("writeTag :: Kepware properties configuration is complete.");

                // read plc tags
                OpcGroupNames.SyncRead((short)OPCAutomation.OPCDataSource.OPCDevice, 1, ref
                   ItemServerHandles, out ItemServerValues, out ItemServerErrors, out yDIR, out zDIR);

                // Add tag value
                ItemServerValues.SetValue(tagValue, 1);

                // Write tag
                OpcGroupNames.SyncWrite(1, ref ItemServerHandles, ref ItemServerValues, out ItemServerErrors);

                //ConnectedOpc.OPCGroups.Remove("AtsWmsCS_6FetchPalletInfoOnLDDetailsGroup");

                return true;

            }
            catch (Exception ex)
            {
                Log.Error("writeTag :: Exception while writing mission data in the plc tag :: " + tagName + " :: " + ex.Message + " stackTrace :: " + ex.StackTrace);
            }

            return false;

        }

        #endregion

        #region Connect and Disconnect PLC

        private void OnConnectPLC()
        {

            Log.Debug("OnConnectPLC :: inside OnConnectPLC");

            try
            {
                // Connection url
                if (!((ConnectedOpc.ServerState.ToString()).Equals("1")))
                {
                    ConnectedOpc.Connect(plcServerConnectionString, "");
                    Log.Debug("OnConnectPLC :: PLC connection successful and OPC server state is :: " + ConnectedOpc.ServerState.ToString());
                }
                else
                {
                    Log.Debug("OnConnectPLC :: Already connected with the plc.");
                }

            }
            catch (Exception ex)
            {
                Log.Error("OnConnectPLC :: Exception while connecting to plc :: " + ex.Message + " stackTrace :: " + ex.StackTrace);
            }
        }

        private void OnDisconnectPLC()
        {
            Log.Debug("inside OnDisconnectPLC");

            try
            {
                ConnectedOpc.Disconnect();
                Log.Debug("OnDisconnectPLC :: Connection with the plc is disconnected.");
            }
            catch (Exception ex)
            {
                Log.Error("OnDisconnectPLC :: Exception while disconnecting to plc :: " + ex.Message);
            }

        }


        #endregion
    }
}


