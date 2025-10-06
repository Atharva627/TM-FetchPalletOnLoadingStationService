using AtsWmsCS_6FetchPalletInfoOnLD.ats_tata_metallics_dbDataSetTableAdapters;
using log4net;
using OPCAutomation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading;
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
        ats_wms_loading_stations_tag_detailsDataTable ats_wms_loading_stations_tag_detailsDataTableStationISActiveDT = null;

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
        OPCGroup CS6FetchPalletOPC;
        //object iDIR;
        // object jDIR;
        object iDIR;
        object jDIR;

        // Connection string
        static string plcServerConnectionString = null;

        #endregion

        #region Global Variables
        static string className = "AtsWmsCS_6FetchPalletInfoOnLDDetails";
        private static readonly ILog Log = LogManager.GetLogger(className);
        private System.Timers.Timer a1FetchPalletInfoOnLDTimer = null;
        string IP_ADDRESS = "";
        int stationId = 0;
        string CS_PalletCode = "";
        string coreQuantity = "";
        string coreMaterialCode = "";
        string palletPresentOnPickup = "";
        string palletCodeOnPickup = "";
        string palletMatrialCodeOnPickup = "";
        float qtyOnPickup;
        string readyToPickup = "";
        int palletStatusId = 0;
        string palletStatusName = "";
        string isActiveValue = "";
        int ccIsActive = 0;
        string RobotRowTag = "";
        string RobotColumnTag = "";
        string RobotLayerTag = "";
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
                    try
                    {
                        ats_wms_master_plc_connection_detailsDataTableDT = ats_wms_master_plc_connection_detailsTableAdapterInstance.GetData();
                        IP_ADDRESS = ats_wms_master_plc_connection_detailsDataTableDT[0].PLC_CONNECTION_IP_ADDRESS;
                        Log.Debug("2.1.1 :: IP_ADDRESS ::" + IP_ADDRESS);
                    }
                    catch (Exception ex)
                    {

                        Log.Error("a1MasterGiveMissionOperation :: Exception Occure while reading machine datasource connection IP_ADDRESS :: " + ex.Message + "StackTrace :: " + ex.StackTrace);
                    }
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


                                int hResult = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
                                string comError = (ex is System.Runtime.InteropServices.COMException) ? ((System.Runtime.InteropServices.COMException)ex).ErrorCode.ToString() : "No COM error";

                                Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while connecting to plc :: " + ex.Message
                                + " HResult :: " + hResult
                                + " COM Component Error :: " + comError
                                + " stackTrace :: " + ex.StackTrace);


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
                                        try
                                        {
                                            if (ats_wms_master_station_detailsDataTableDT != null && ats_wms_master_station_detailsDataTableDT.Count > 0)

                                            {
                                                for (int i = 0; i < ats_wms_master_station_detailsDataTableDT.Count; i++)
                                                {
                                                    Thread.Sleep(100);
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
                                                            //Checking Is Active in CC
                                                            ccIsActive = ats_wms_loading_stations_tag_detailsDataTableDT[0].STATION_IS_ACTIVE;
                                                            if (ccIsActive == 1)
                                                            {


                                                                Log.Debug("7 :: The Station is Active");
                                                                // Checking is active in Tag
                                                                isActiveValue = readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].IS_ACTIVE_TAG);
                                                                try
                                                                {
                                                                    if (isActiveValue.Equals("True"))

                                                                    {
                                                                        try
                                                                        {
                                                                            //checking if the pallet is present
                                                                            palletPresentOnPickup = readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].PICKUP_POSITION_PALLET_PRESENT_TAG);

                                                                            if (palletPresentOnPickup.Equals("True"))

                                                                            {
                                                                                Log.Debug("8 :: Pallet is present at Loading Station");
                                                                                //reading pallet code
                                                                                try
                                                                                {
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

                                                                                        Log.Debug("9 :: Reading Pallet Code and Checking its Format");

                                                                                        //checking if the pallet information details are already avalilable if no inserting the new one
                                                                                        ats_wms_master_pallet_informationDataTableDT = ats_wms_master_pallet_informationTableAdapterInstance.GetDataByPALLET_CODEOrderByPALLET_INFORMATION_IDDESC(palletCodeOnPickup);

                                                                                        if (ats_wms_master_pallet_informationDataTableDT != null && ats_wms_master_pallet_informationDataTableDT.Count > 0)
                                                                                        {
                                                                                            if (ats_wms_master_pallet_informationDataTableDT[0].LOADING_STATION_WORKDONE == 0)
                                                                                            {

                                                                                            
                                                                                            //update details in the station tag tables 
                                                                                            //ready material code & QTY from PLC tag
                                                                                            //check if pichup is ready
                                                                                            Log.Debug("10 :: Checking for Data entry in DB for the Pallet Code");

                                                                                            //checking if the pallet is ready for pickup
                                                                                            try
                                                                                            {
                                                                                                readyToPickup = readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].PICKUP_IS_READY_TAG);
                                                                                            }
                                                                                            catch (Exception ex)
                                                                                            {

                                                                                                int hResult = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
                                                                                                string comError = (ex is System.Runtime.InteropServices.COMException) ? ((System.Runtime.InteropServices.COMException)ex).ErrorCode.ToString() : "No COM error";

                                                                                                Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while reading readyToPickup tag :: " + ex.Message
                                                                                                + " HResult :: " + hResult
                                                                                                + " COM Component Error :: " + comError
                                                                                                + " stackTrace :: " + ex.StackTrace);
                                                                                            }

                                                                                            if (readyToPickup.Equals("True"))
                                                                                            {
                                                                                                for (;;)
                                                                                                {
                                                                                                    Thread.Sleep(1000);
                                                                                                    Log.Debug("11 :: Reading Pallet Code at Pickup Position");
                                                                                                    //getting pallet material and qty details from PLC
                                                                                                    try
                                                                                                    {
                                                                                                        palletMatrialCodeOnPickup = readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].PICKUP_POSITION_MATERIAL_CODE_TAG).Trim();
                                                                                                    }
                                                                                                    catch (Exception ex)
                                                                                                    {

                                                                                                        int hResult = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
                                                                                                        string comError = (ex is System.Runtime.InteropServices.COMException) ? ((System.Runtime.InteropServices.COMException)ex).ErrorCode.ToString() : "No COM error";

                                                                                                        Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while reading palletMatrialCodeOnPickup tag :: " + ex.Message
                                                                                                        + " HResult :: " + hResult
                                                                                                        + " COM Component Error :: " + comError
                                                                                                        + " stackTrace :: " + ex.StackTrace);
                                                                                                    }

                                                                                                    Log.Debug("12 :: Reading Pickup Position Quantity");

                                                                                                    try
                                                                                                    {
                                                                                                        qtyOnPickup = float.Parse(readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].PICKUP_POSITION_QUANTITY_VALUE_TAG));
                                                                                                    }
                                                                                                    catch (Exception ex)
                                                                                                    {

                                                                                                        int hResult = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
                                                                                                        string comError = (ex is System.Runtime.InteropServices.COMException) ? ((System.Runtime.InteropServices.COMException)ex).ErrorCode.ToString() : "No COM error";

                                                                                                        Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while reading qtyOnPickup tag :: " + ex.Message
                                                                                                        + " HResult :: " + hResult
                                                                                                        + " COM Component Error :: " + comError
                                                                                                        + " stackTrace :: " + ex.StackTrace);

                                                                                                    }
                                                                                                    if (palletMatrialCodeOnPickup != null && palletMatrialCodeOnPickup == "EMPTY")
                                                                                                    {

                                                                                                        Log.Debug("12.1  :: Found  palletMatrialCodeOnPickup  Empty :: set it as a Empty pallet");

                                                                                                        ats_wms_master_pallet_informationDataTableDT = ats_wms_master_pallet_informationTableAdapterInstance.GetDataByPALLET_CODEOrderByPALLET_INFORMATION_IDDESC(palletCodeOnPickup);

                                                                                                        if (ats_wms_master_pallet_informationDataTableDT != null && ats_wms_master_pallet_informationDataTableDT.Count > 0)
                                                                                                        {

                                                                                                            try
                                                                                                            {
                                                                                                                Log.Debug("2.1 ::Updating Rework/Reject station pallet details");

                                                                                                                if (ats_wms_master_pallet_informationDataTableDT[0].LOADING_STATION_WORKDONE == 0)
                                                                                                                {


                                                                                                                    //update in pallet information table
                                                                                                                    ats_wms_master_pallet_informationTableAdapterInstance.UpdatePalletInfoDetails(0,
                                                                                                                        "NA",
                                                                                                                        palletMatrialCodeOnPickup, 0, "NA", "NA", "NA",
                                                                                                                        ats_wms_master_station_detailsDataTableDT[i].STATION_ID,
                                                                                                                        0,
                                                                                                                        3,
                                                                                                                        "EMPTY",
                                                                                                                        0,
                                                                                                                        0,
                                                                                                                        DateTime.Now.ToString(),
                                                                                                                        "NA",
                                                                                                                        1,
                                                                                                                        0, ats_wms_master_station_detailsDataTableDT[i].STATION_NAME,
                                                                                                                         0, 0, 0,
                                                                                                                        ats_wms_master_station_detailsDataTableDT[i].PALLET_TYPE_ID,
                                                                                                                        ats_wms_master_pallet_informationDataTableDT[0].PALLET_INFORMATION_ID

                                                                                                                        );
                                                                                                                    Log.Debug("2.2 :: Pallet is EMPTY update  Successfully");
                                                                                                                    break;
                                                                                                                }



                                                                                                            }
                                                                                                            catch (Exception ex)
                                                                                                            {
                                                                                                                int hResult = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
                                                                                                                string comError = (ex is System.Runtime.InteropServices.COMException) ? ((System.Runtime.InteropServices.COMException)ex).ErrorCode.ToString() : "No COM error";

                                                                                                                Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while updating pallet data :: " + ex.Message
                                                                                                                + " HResult :: " + hResult
                                                                                                                + " COM Component Error :: " + comError
                                                                                                                + " stackTrace :: " + ex.StackTrace);

                                                                                                            }
                                                                                                        }

                                                                                                    }

                                                                                                    //checking if the values are not empty
                                                                                                    else if (palletMatrialCodeOnPickup != null && palletMatrialCodeOnPickup != "" && palletMatrialCodeOnPickup != "0" && qtyOnPickup > 0.0f)
                                                                                                    {

                                                                                                        try
                                                                                                        {
                                                                                                            RobotRowTag = readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].ROBOT_ROW_TAG);
                                                                                                        }
                                                                                                        catch (Exception ex)
                                                                                                        {

                                                                                                            int hResult = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
                                                                                                            string comError = (ex is System.Runtime.InteropServices.COMException) ? ((System.Runtime.InteropServices.COMException)ex).ErrorCode.ToString() : "No COM error";

                                                                                                            Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while reading RobotRowTag tag :: " + ex.Message
                                                                                                            + " HResult :: " + hResult
                                                                                                            + " COM Component Error :: " + comError
                                                                                                            + " stackTrace :: " + ex.StackTrace);
                                                                                                        }

                                                                                                        if (RobotRowTag != null && RobotRowTag != "" && RobotRowTag != "0")
                                                                                                        {
                                                                                                            Log.Debug("9.1 :: Reading ROBOT_ROW_TAG");

                                                                                                            try
                                                                                                            {
                                                                                                                RobotColumnTag = readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].ROBOT_COLUMN_TAG);
                                                                                                            }
                                                                                                            catch (Exception ex)
                                                                                                            {
                                                                                                                int hResult = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
                                                                                                                string comError = (ex is System.Runtime.InteropServices.COMException) ? ((System.Runtime.InteropServices.COMException)ex).ErrorCode.ToString() : "No COM error";

                                                                                                                Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while reading RobotColumnTag tag :: " + ex.Message
                                                                                                                + " HResult :: " + hResult
                                                                                                                + " COM Component Error :: " + comError
                                                                                                                + " stackTrace :: " + ex.StackTrace);

                                                                                                            }

                                                                                                            if (RobotColumnTag != null && RobotColumnTag != "" && RobotColumnTag != "0")
                                                                                                            {
                                                                                                                Log.Debug("9.2 :: Reading ROBOT_COLUMN_TAG");

                                                                                                                try
                                                                                                                {
                                                                                                                    RobotLayerTag = readTag(ats_wms_loading_stations_tag_detailsDataTableDT[0].ROBOT_LAYER_TAG);
                                                                                                                }
                                                                                                                catch (Exception ex)
                                                                                                                {

                                                                                                                    int hResult = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
                                                                                                                    string comError = (ex is System.Runtime.InteropServices.COMException) ? ((System.Runtime.InteropServices.COMException)ex).ErrorCode.ToString() : "No COM error";

                                                                                                                    Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while reading robotlayer tag :: " + ex.Message
                                                                                                                    + " HResult :: " + hResult
                                                                                                                    + " COM Component Error :: " + comError
                                                                                                                    + " stackTrace :: " + ex.StackTrace);
                                                                                                                }

                                                                                                                if (RobotLayerTag != null && RobotLayerTag != "" && RobotLayerTag != "0")
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
                                                                                                                                ats_wms_loading_stations_tag_detailsTableAdapterInstance.UpdateStationPalletDetails(1, palletCodeOnPickup, palletMatrialCodeOnPickup, qtyOnPickup, Convert.ToInt32(RobotRowTag), Convert.ToInt32(RobotColumnTag), Convert.ToInt32(RobotLayerTag), ats_wms_loading_stations_tag_detailsDataTableDT[0].TAG_DETAILS_ID);
                                                                                                                                Log.Debug("15.1 :: Updated new Pallet information in Loading Station");
                                                                                                                            }
                                                                                                                        }

                                                                                                                        Log.Debug("15.2 :: Updating new Pallet information in Loading Station");
                                                                                                                        //fetching pallet info for checking if the pallet information is already updated or not
                                                                                                                        ats_wms_master_pallet_informationDataTableDT = ats_wms_master_pallet_informationTableAdapterInstance.GetDataByPALLET_CODEOrderByPALLET_INFORMATION_IDDESC(palletCodeOnPickup);
                                                                                                                        if (ats_wms_master_pallet_informationDataTableDT != null && ats_wms_master_pallet_informationDataTableDT.Count > 0)
                                                                                                                        {

                                                                                                                            Log.Debug("16 :: FEtching Pallet info for the Pallet Code ");

                                                                                                                            if (ats_wms_master_pallet_informationDataTableDT[0].LOADING_STATION_WORKDONE == 0)
                                                                                                                            {


                                                                                                                                int reworkStatinPalletId = ats_wms_master_pallet_informationDataTableDT[0].PALLET_TYPE_ID;

                                                                                                                                Log.Debug("17 :: Updating information against the Pallet Code");

                                                                                                                                try
                                                                                                                                {
                                                                                                                                    if (ats_wms_master_station_detailsDataTableDT[i].STATION_NAME == "REWORK" || ats_wms_master_station_detailsDataTableDT[i].STATION_NAME == "REJECT")
                                                                                                                                    {
                                                                                                                                        //update in pallet information table
                                                                                                                                        ats_wms_master_pallet_informationTableAdapterInstance.UpdatePalletInfoDetails(ats_wms_master_product_detailsDataTableDT[0].PRODUCT_ID,
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
                                                                                                                                             Convert.ToInt32(RobotRowTag), Convert.ToInt32(RobotColumnTag), Convert.ToInt32(RobotLayerTag),
                                                                                                                                            reworkStatinPalletId,
                                                                                                                                            ats_wms_master_pallet_informationDataTableDT[0].PALLET_INFORMATION_ID

                                                                                                                                            );
                                                                                                                                        Log.Debug("18 :: Pallet is Loaded Successfully");
                                                                                                                                        break;

                                                                                                                                    }
                                                                                                                                    else
                                                                                                                                    {
                                                                                                                                        //update in pallet information table
                                                                                                                                        ats_wms_master_pallet_informationTableAdapterInstance.UpdatePalletInfoDetails(ats_wms_master_product_detailsDataTableDT[0].PRODUCT_ID,
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
                                                                                                                                             Convert.ToInt32(RobotRowTag), Convert.ToInt32(RobotColumnTag), Convert.ToInt32(RobotLayerTag),
                                                                                                                                            ats_wms_master_station_detailsDataTableDT[i].PALLET_TYPE_ID,
                                                                                                                                            ats_wms_master_pallet_informationDataTableDT[0].PALLET_INFORMATION_ID

                                                                                                                                            );
                                                                                                                                        Log.Debug("18 :: Pallet is Loaded Successfully");
                                                                                                                                        break;

                                                                                                                                    }

                                                                                                                                }
                                                                                                                                catch (Exception ex)
                                                                                                                                {

                                                                                                                                    int hResult = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
                                                                                                                                    string comError = (ex is System.Runtime.InteropServices.COMException) ? ((System.Runtime.InteropServices.COMException)ex).ErrorCode.ToString() : "No COM error";

                                                                                                                                    Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while updating loaded pallet data :: " + ex.Message
                                                                                                                                    + " HResult :: " + hResult
                                                                                                                                    + " COM Component Error :: " + comError
                                                                                                                                    + " stackTrace :: " + ex.StackTrace);
                                                                                                                                }

                                                                                                                            }
                                                                                                                            break;

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
                                                                                                }
                                                                                            }
                                                                                        }
                                                                                            else
                                                                                            {
                                                                                                Log.Debug(" :: loading station workdone is not a zero :: palletcode:: " + palletCodeOnPickup);
                                                                                                break;
                                                                                            }


                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            try
                                                                                            {
                                                                                                Log.Debug("2.1 :: Pallet Details not present in the DB.... Inserting New Information");
                                                                                                //pallet details are not into database already inserting new 
                                                                                                ats_wms_master_pallet_informationTableAdapterInstance.Insert(palletCodeOnPickup, 0, "NA", "0", 0, "NA", "NA", "NA", 0,
                                                                                                   "NA", 0, 3, "Empty", 0, 0, DateTime.Now, 1, "master_user", 0, "0", 0, "NA", 0, 0, 0, 0, 0);
                                                                                                break;
                                                                                            }
                                                                                            catch (Exception ex)
                                                                                            {
                                                                                                int hResult = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
                                                                                                string comError = (ex is System.Runtime.InteropServices.COMException) ? ((System.Runtime.InteropServices.COMException)ex).ErrorCode.ToString() : "No COM error";

                                                                                                Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while updating pallet data :: " + ex.Message
                                                                                                + " HResult :: " + hResult
                                                                                                + " COM Component Error :: " + comError
                                                                                                + " stackTrace :: " + ex.StackTrace);

                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }
                                                                                catch (Exception ex)
                                                                                {

                                                                                    int hResult = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
                                                                                    string comError = (ex is System.Runtime.InteropServices.COMException) ? ((System.Runtime.InteropServices.COMException)ex).ErrorCode.ToString() : "No COM error";

                                                                                    Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while checking pallet code on pick up :: " + ex.Message
                                                                                    + " HResult :: " + hResult
                                                                                    + " COM Component Error :: " + comError
                                                                                    + " stackTrace :: " + ex.StackTrace);
                                                                                }
                                                                            }

                                                                            else
                                                                            {
                                                                                Log.Debug("15.3 :: Pallet Present is False... reseting all the Values ");
                                                                                //as pallet present is false reseting value to 0 and NA after checking if value are not already reset 
                                                                                ats_wms_loading_stations_tag_detailsDataTableDT = ats_wms_loading_stations_tag_detailsTableAdapterInstance.GetDataBySTATION_ID(ats_wms_master_station_detailsDataTableDT[i].STATION_ID);
                                                                                if (ats_wms_loading_stations_tag_detailsDataTableDT != null && ats_wms_loading_stations_tag_detailsDataTableDT.Count > 0)
                                                                                {
                                                                                    if (ats_wms_loading_stations_tag_detailsDataTableDT[0].DROP_POSITION_PALLET_CODE_VALUE != "NA")
                                                                                    {
                                                                                        ats_wms_loading_stations_tag_detailsTableAdapterInstance.UpdateResetData(0, 0, "NA", "NA", 0, 0, 0, 0, ats_wms_loading_stations_tag_detailsDataTableDT[0].TAG_DETAILS_ID);
                                                                                        Log.Debug("15.3.1 :: Values Reset");
                                                                                        break;
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                        catch (Exception ex)
                                                                        {

                                                                            int hResult = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
                                                                            string comError = (ex is System.Runtime.InteropServices.COMException) ? ((System.Runtime.InteropServices.COMException)ex).ErrorCode.ToString() : "No COM error";

                                                                            Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while checking pallet present :: " + ex.Message
                                                                            + " HResult :: " + hResult
                                                                            + " COM Component Error :: " + comError
                                                                            + " stackTrace :: " + ex.StackTrace);
                                                                        }
                                                                    }

                                                                    else
                                                                    {
                                                                        Log.Debug("7.2 :: The Station is Not Active in CC");
                                                                    }
                                                                }
                                                                catch (Exception ex)
                                                                {

                                                                    int hResult = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
                                                                    string comError = (ex is System.Runtime.InteropServices.COMException) ? ((System.Runtime.InteropServices.COMException)ex).ErrorCode.ToString() : "No COM error";

                                                                    Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while checking pallet present :: " + ex.Message
                                                                    + " HResult :: " + hResult
                                                                    + " COM Component Error :: " + comError
                                                                    + " stackTrace :: " + ex.StackTrace);
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
                                        catch (Exception ex)
                                        {

                                            int hResult = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
                                            string comError = (ex is System.Runtime.InteropServices.COMException) ? ((System.Runtime.InteropServices.COMException)ex).ErrorCode.ToString() : "No COM error";

                                            Log.Error("atsWmsCS_6FetchPalletInfoOnLDDetailsOperation :: Exception while connecting to plc :: " + ex.Message
                                            + " HResult :: " + hResult
                                            + " COM Component Error :: " + comError
                                            + " stackTrace :: " + ex.StackTrace);
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
                CS6FetchPalletOPC = ConnectedOpc.OPCGroups.Add("AtsWmsCS_6FetchPalletInfoOnLDDetailsGroup");
                CS6FetchPalletOPC.DeadBand = 0;
                CS6FetchPalletOPC.UpdateRate = 100;
                CS6FetchPalletOPC.IsSubscribed = true;
                CS6FetchPalletOPC.IsActive = true;
                CS6FetchPalletOPC.OPCItems.AddItems(1, ref OPCItemIDs, ref ClientHandles, out ItemServerHandles, out ItemServerErrors, RequestedDataTypes, AccessPaths);
                //Log.Debug("readTag :: Kepware properties configuration is complete.");

                // Read tag
                CS6FetchPalletOPC.SyncRead((short)OPCAutomation.OPCDataSource.OPCDevice, 1, ref
                   ItemServerHandles, out ItemServerValues, out ItemServerErrors, out iDIR, out jDIR);

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
                CS6FetchPalletOPC = ConnectedOpc.OPCGroups.Add("AtsWmsCS_6FetchPalletInfoOnLDDetailsGroup");
                CS6FetchPalletOPC.DeadBand = 0;
                CS6FetchPalletOPC.UpdateRate = 100;
                CS6FetchPalletOPC.IsSubscribed = true;
                CS6FetchPalletOPC.IsActive = true;
                CS6FetchPalletOPC.OPCItems.AddItems(1, ref OPCItemIDs, ref ClientHandles, out ItemServerHandles, out ItemServerErrors, RequestedDataTypes, AccessPaths);
                //Log.Debug("writeTag :: Kepware properties configuration is complete.");

                // read plc tags
                CS6FetchPalletOPC.SyncRead((short)OPCAutomation.OPCDataSource.OPCDevice, 1, ref
                   ItemServerHandles, out ItemServerValues, out ItemServerErrors, out iDIR, out jDIR);

                // Add tag value
                ItemServerValues.SetValue(tagValue, 1);

                // Write tag
                CS6FetchPalletOPC.SyncWrite(1, ref ItemServerHandles, ref ItemServerValues, out ItemServerErrors);

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




