using INX_AGVC;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace INX
{
    public partial class Inventory : Form
    {
        public Inventory()
        {
            InitializeComponent();
            Global.ReadINI();
        }

        ucShelf[] ucShelf;
        ClsExport objExp = new ClsExport();
        private string strIP1 = string.Empty;
        private string strIP2 = string.Empty;
        
        string strPresent = string.Empty;
        ToolTip tooltipADAM = new ToolTip();
        DataTable dtResult = new DataTable();

        int Match_Type_Present = 1;
        int Match_Type_ItemID = 2;
        int Match_Type_Storage = 3;
        int Match_Type_SAME = 0;


        private void ucShelf_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ucShelf ucShelf = (ucShelf)sender;
               
            }
        }
        private void ucShelf_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ucShelf ucShelf = (ucShelf)sender;
                txtShelfID.Text = ucShelf.ShelfID;
                txtItemDetail.Text = ucShelf.ItemDetail;
                txtMemo.Text = ucShelf.Memo;
                if (!string.IsNullOrEmpty(ucShelf.ItemID_Cycle))
                {
                    lblCycleID.Text = ucShelf.ItemID_Cycle;
                    this.lblCycleID.Font = new Font(FontFamily.GenericSansSerif,
           15.0F, FontStyle.Bold);
                    lblCycleID.ForeColor = Color.Red;
                }
                
                txtItemID.Text = ucShelf.ItemID;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            funUpdateUnLock();
            this.Close();
        }

        private void funUpdateUnLock()
        {
            string strCmd = string.Empty;
            bool blnOK = false;

            strCmd = "UPDATE ERACK_INFO SET Status = 'ENABLE' WHERE ErackID = '" + cboErackID.Text + "'";

            ERACK_EVENT_HISTORY erack_event_history = new ERACK_EVENT_HISTORY();
            erack_event_history.LINE = Global.g_INI.Common.LINE;
            erack_event_history.FLOOR = Global.g_INI.Common.FLOOR;
            erack_event_history.SystemID = Global.g_INI.Common.SystemID;
            erack_event_history.ErackType = "ERACK";
            erack_event_history.ErackID = cboErackID.Text;
            erack_event_history.EventType = "ENABLE";
            erack_event_history.ACK = "Y";
            erack_event_history.CreateTime = ClsGlobalFunc.GetFormatDT();
            erack_event_history.Memo = "[ENABLE] Lock Erack For Inventory";

            try
            {
                blnOK = ClsMSSQL.SqlNonQuery(strCmd, Global.g_INI.DB.ConnStr);
                if (blnOK && funInsertErackEventHistory(erack_event_history))
                {
                    lblMSG.Text = "[Erack]" + cboErackID.Text + " is UNLOCK !";
                    btnLock.Image = Properties.Resources.unlock;
                }
                else
                    return;
            }
            catch (Exception ex)
            {
                ClsLog.TraceLog(ex.Message, EnuLogType.Exception);
                return;
            }
        }

        private void Cycle_Load(object sender, EventArgs e)
        {
            //初始EventType 下拉選單
            cboEventType.Items.Add("LOAD");
            cboEventType.Items.Add("CLEAR");  //INITIAL -> CLEAR
            //初始ErackID 下拉選單
            string strCmd = string.Empty;
            bool blnOK = false;
            DataTable dtTmp = new DataTable();
            try
            {
                strCmd = "SELECT * FROM ERACK_INFO ORDER BY ErackID ASC";
                blnOK = ClsMSSQL.GetDBData(strCmd, Global.g_INI.DB.ConnStr, ref dtTmp);
                if (blnOK)
                {
                    if (dtTmp.Rows.Count > 0)
                    {
                        for (int i = 0; i < dtTmp.Rows.Count; i++)
                            cboErackID.Items.Add(dtTmp.Rows[i]["ErackID"].ToString());
                    }
                }
                if (Global.g_INI.Common.PassHostTrx)
                {
                    lblItemDetail.Visible = false;
                    txtItemDetail.Visible = false;
                }
                else
                {
                    lblItemDetail.Visible = true;
                    txtItemDetail.Visible = true;
                }

                if (Global.g_INI.Common.PassBCRCheck)
                    btnGetID.Visible = false;
                else
                    btnGetID.Visible = true;
            }
            catch (Exception ex)
            {
                ClsLog.TraceLog(ex.ToString(), EnuLogType.Exception);
                return;
            }
            finally
            {
                if (dtTmp != null)
                    dtTmp.Dispose();
            }
            cboErackID.SelectedIndex = 0;

            if (funGetErackStatus() == "DISABLE")
                btnLock.Image = Properties.Resources._lock;
            else if (funGetErackStatus() == "ENABLE")
                btnLock.Image = Properties.Resources.unlock;
            funLoadingShelfInfo(cboErackID.Text);
            //rbt_Manual.Checked = true;
            if (Global.g_INI.Common.PassBCRCheck)
                btnGetID.Visible = false;
        }
        //private void rbt_samiAuto_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (rbt_samiAuto.Checked)
        //    {
        //        btn_InventoryResult.Enabled = true;
        //        btn_Import.Enabled = true;
        //        if (panel_ShelfInfo != null)
        //        {
        //            panel_ShelfInfo.Controls.Clear();
        //        }
        //    }
        //}

        //private void rbt_Manual_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (rbt_Manual.Checked)
        //    {
        //        btn_InventoryResult.Enabled = false;
        //        btn_Import.Enabled = false;
        //        if (!string.IsNullOrEmpty(cboErackID.Text))
        //            funLoadingShelfInfo(cboErackID.Text);
        //        else
        //        {
        //            MessageBox.Show("Please Choose ErackID First !!", "Hint", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //            return;
        //        }
        //    }
        //}

        private void Create_InventoryResult()
        {
            string strCmd = string.Empty;
            bool blnOK = false;
            DataTable dtTMP = new DataTable();
            int j = 0;
            string strCycleID = string.Empty;

            try
            {
                //if (!rbt_samiAuto.Checked)
                //{
                //    MessageBox.Show("It only support Semi-Automatic Inventory" , "Hint", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    return;
                //}
                if (cboErackID.SelectedIndex < 0)
                {
                    MessageBox.Show("Please Choose ErackID", "Hint", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                //開始比對前要先把Erack鎖住，避免預約
                funUpdateLock();

                //開始比對
                strCmd = "SELECT * FROM SHELF_INFO WHERE ErackID ='" + cboErackID.Text + "'";
                blnOK = ClsMSSQL.GetDBData(strCmd, Global.g_INI.DB.ConnStr, ref dtTMP);
                if (blnOK && dtTMP.Rows.Count > 0)
                {
                    for (int i = 0; i < dtTMP.Rows.Count; i++)
                    {
                        if (dtResult.Rows[j]["ShelfID"].ToString() == dtTMP.Rows[i]["ShelfID"].ToString())  //再匯入的表跟資料庫撈出來的表比對兩個儲位ID
                        {
                            if (dtResult.Rows[j]["OccupyID"].ToString().Trim() != dtTMP.Rows[i]["OccupyID"].ToString().Trim())
                            {
                                strCycleID = dtResult.Rows[i]["OccupyID"].ToString();
                                funUpdateMatchType(dtTMP.Rows[j]["ShelfID"].ToString(), Match_Type_ItemID, strCycleID);
                                j++;
                            }
                            else
                            {
                                strCycleID = " ";
                                funUpdateMatchType(dtTMP.Rows[j]["ShelfID"].ToString(), Match_Type_SAME, strCycleID);
                                j++;
                            }
                        }
                    }
                }

                //MessageBox.Show("Please Input Inventory Result FileName, First !!", "Hint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //int intFileType = 1;
                //string sFileName = objExp.funSaveFile("InventoryResult_" + DateTime.Today.ToString("yyyyMMdd") + "_" + cboErackID.Text, out intFileType);

                //if (string.IsNullOrWhiteSpace(sFileName))
                //{
                //    MessageBox.Show("Please Input the FileName First", "Hint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    return;
                //}

                //ClsExport.funExpExcel(sFileName, dtTMP);

            }
            catch (Exception ex)
            {
                ClsLog.TraceLog(ex.ToString(), EnuLogType.Exception);
                return;
            }
            finally
            {
                if (dtTMP != null)
                    dtTMP.Dispose();
                if (dtResult != null)
                    dtResult.Dispose();
            }

            if (!funLoadingInventoryReuslt())
            {
                lblMSG.Text = "Loading the Result of Inventory Data FAIL!!";
                return;
            }
            else
                lblMSG.Text = "Loading the Result of Inventory Data SUCCESS!!";

            if (!funCheckInventoryResult())
            {
                MessageBox.Show("There is no data need to adjust");
                funUpdateUnLock();
            }
        }

        private void funUpdateLock()
        {
            string strCmd = string.Empty;
            bool blnOK = false;

            strCmd = "UPDATE ERACK_INFO SET Status = 'DISABLE' WHERE ErackID = '"+cboErackID.Text+"'";

            ERACK_EVENT_HISTORY erack_event_history = new ERACK_EVENT_HISTORY();
            erack_event_history.LINE = Global.g_INI.Common.LINE;
            erack_event_history.FLOOR = Global.g_INI.Common.FLOOR;
            erack_event_history.SystemID = Global.g_INI.Common.SystemID;
            erack_event_history.ErackType = "ERACK";
            erack_event_history.ErackID = cboErackID.Text;
            erack_event_history.EventType = "DISABLE";
            erack_event_history.ACK = "Y";
            erack_event_history.CreateTime = ClsGlobalFunc.GetFormatDT();
            erack_event_history.Memo = "[DISABLE] Lock Erack For Inventory";
          
            try
            {
                blnOK = ClsMSSQL.SqlNonQuery(strCmd, Global.g_INI.DB.ConnStr);
                if (blnOK && funInsertErackEventHistory(erack_event_history))
                {
                    lblMSG.Text = "[Erack]" + cboErackID.Text + " is LOCK !";
                    btnLock.Image = Properties.Resources._lock;
                }
                else
                    return;
            }
            catch (Exception ex)
            {
                ClsLog.TraceLog(ex.Message, EnuLogType.Exception);
                return;
            }
        }

        private bool funCheckInventoryResult()
        {
            string strCmd = string.Empty;
            bool blnOK = false;
            DataTable dtTmp = new DataTable();
            bool Result = false;

            strCmd = "SELECT Count(*) as CNT_MatchType FROM SHELF_INFO WHERE ErackID  ='"+cboErackID.Text+"' AND Match_Type='2'";
            blnOK = ClsMSSQL.GetDBData(strCmd, Global.g_INI.DB.ConnStr, ref dtTmp,true);
            if (blnOK)
            {
                if (Convert.ToInt32(dtTmp.Rows[0]["CNT_MatchType"].ToString())>0)
                    Result= true;
            }
            return Result;
        }

        private void funUpdateMatchType(string strShelfID, int MatchResultType, string str_CycleID)
        {
            string strCMD = "UPDATE SHELF_INFO SET Match_Type = '" + MatchResultType +"' , CycleID = '"+str_CycleID+"' WHERE ShelfID = '"+ strShelfID + "'";
            bool blnOK = ClsMSSQL.SqlNonQuery(strCMD, Global.g_INI.DB.ConnStr);
            
        }

        private bool funLoadingInventoryReuslt() 
        {
            string strCmd = string.Empty;
            bool blnOK = false;
            DataTable dtTmp = new DataTable();
            bool flag = false;
            try
            {
                strCmd = "SELECT * FROM SHELF_INFO WHERE ErackID = '"+cboErackID.Text+"'";
                blnOK = ClsMSSQL.GetDBData(strCmd, Global.g_INI.DB.ConnStr, ref dtTmp);

                if (blnOK)
                {
                    if (dtTmp.Rows.Count > 0)
                    {

                        if (panel_ShelfInfo != null)
                            panel_ShelfInfo.Controls.Clear();
                        //載入格位排版
                        ucShelf = new ucShelf[12];
                        for (int j = 0; j < ucShelf.Length; j++)
                        {
                            ucShelf[j] = new ucShelf();
                            panel_ShelfInfo.Controls.Add(ucShelf[j]);
                            ucShelf[j].Width = (this.Width - 450) / 4;
                            ucShelf[j].Height = (int)(ucShelf[j].Width * 0.9);
                            ucShelf[j].Top = 0 + (j / 4) * ucShelf[j].Height;
                            ucShelf[j].Left = 3 + (j % 4) * ucShelf[j].Width;
                            ucShelf[j].Visible = false;
                            ucShelf[j].Tag = j;
                            ucShelf[j].MouseClick += new System.Windows.Forms.MouseEventHandler(ucShelf_MouseClick);

                        }
                        int intCol = dtTmp.Rows.Count / 3;
                        for (int i = 0; i < dtTmp.Rows.Count; i++)
                        {
                            int intIdx = (i % intCol) + (i / intCol) * 4;
                            ucShelf[intIdx].ShelfID = dtTmp.Rows[i]["ShelfID"].ToString().Trim();
                            ucShelf[intIdx].Visible = true;


                            string strOccupyID = string.IsNullOrEmpty(dtTmp.Rows[i]["OccupyID"].ToString().Trim()) ? string.Empty : dtTmp.Rows[i]["OccupyID"].ToString().Trim();
                            string strReservedID = string.IsNullOrEmpty(dtTmp.Rows[i]["ReservedID"].ToString().Trim()) ? string.Empty : dtTmp.Rows[i]["ReservedID"].ToString().Trim();
                            string strStatus = string.IsNullOrEmpty(dtTmp.Rows[i]["Status"].ToString().Trim()) ? string.Empty : dtTmp.Rows[i]["Status"].ToString().Trim();
                            bool blnDisable = string.IsNullOrEmpty(dtTmp.Rows[i]["Enable"].ToString().Trim()) ? false : (dtTmp.Rows[i]["Enable"].ToString().Trim() == "N");
                            string strPresent = string.IsNullOrEmpty(dtTmp.Rows[i]["Present"].ToString().Trim()) ? string.Empty : dtTmp.Rows[i]["Present"].ToString().Trim();
                            bool blnNGPort = string.IsNullOrEmpty(dtTmp.Rows[i]["Purpose"].ToString().Trim()) ? false : (dtTmp.Rows[i]["Purpose"].ToString().Trim() == "N");
                            string strPurpose = string.IsNullOrEmpty(dtTmp.Rows[i]["Purpose"].ToString().Trim()) ? "G" : dtTmp.Rows[i]["Purpose"].ToString().Trim();
                            string strManual = string.IsNullOrEmpty(dtTmp.Rows[i]["Manual"].ToString().Trim()) ? string.Empty : (dtTmp.Rows[i]["Manual"].ToString().Trim());


                            if (blnDisable)  //禁用
                            {
                                ucShelf[intIdx].ShelfDisable(strOccupyID);
                            }
                            else if (strStatus == "ALARM")  //異常
                            {
                                ucShelf[intIdx].ShelfAlarm(strOccupyID);
                            }
                            else   //填入儲位狀態和顯示FoupID
                            {
                                if (!string.IsNullOrEmpty(strOccupyID))
                                {
                                    ucShelf[intIdx].ShelfLoad(strOccupyID);
                                    //ucShelf[intIdx].ManualChange(strManual);  //顯示自動或手動
                                   
                                }
                                else if (!string.IsNullOrEmpty(strReservedID))
                                {
                                    ucShelf[intIdx].ShelfReserved(strReservedID);
                                }
                                else
                                {
                                    ucShelf[intIdx].ShelfUnLoad("");
                                    //ucShelf[intIdx].ManualChange(strManual);
                                }
                            }

                            if (!string.IsNullOrEmpty(dtTmp.Rows[i]["Match_Type"].ToString())) /*&& !string.IsNullOrEmpty(dtTmp.Rows[i]["CycleID"].ToString()))*/
                            {
                                if (Convert.ToInt32(dtTmp.Rows[i]["Match_Type"].ToString()) == Match_Type_ItemID
                                && dtTmp.Rows[i]["OccupyID"].ToString() != dtTmp.Rows[i]["CycleID"].ToString())
                                {
                                    //if (strStatus == "FULL" || strStatus == "EMPTY")  //只看FULL & EMPTY , 預約跟其他狀態不看
                                    //{
                                    ucShelf[intIdx].ItemID_Cycle = string.IsNullOrEmpty(dtTmp.Rows[i]["CycleID"].ToString()) ? string.Empty : dtTmp.Rows[i]["CycleID"].ToString().Trim();
                                    ucShelf[intIdx].ShelfCycleDiff(ucShelf[intIdx].ShelfID, ucShelf[intIdx].ItemID_Cycle, strOccupyID);

                                    // }
                                }
                            }


                        }
                        flag = true;
                    }
                }

            }
            catch (Exception ex)
            {
                ClsLog.TraceLog(ex.Message, EnuLogType.Exception);
            }
            return flag;

        }


        private void btn_Import_Click(object sender, EventArgs e)
        {
            DataTable dtTmp = new DataTable();
            DataTable dtTmp1 = new DataTable();
            string delimiter = "|";
            string strEM = string.Empty;
            string strSql = string.Empty;

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "XLS檔(*.xls)|*.xlsx|所有檔案(*.*)|*.*";
            dialog.InitialDirectory = "C:";
            dialog.Title = "選擇匯入檔案";
            string strFilePath = string.Empty;
            string strFileExt = string.Empty;

            //if (!rbt_samiAuto.Checked)
            //{
            //    MessageBox.Show("Please Choose Semi-Automatic Inventory Mode!!", "Hint", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    return;
            //}

            if (cboErackID.SelectedIndex<0)
            {
                MessageBox.Show("Please Choose ErackID First", "Hint", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            else
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    //this.DialogResult = System.Windows.Forms.MessageBox.Show("Start to Import the Data！", "Hint", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    //if (this.DialogResult == System.Windows.Forms.DialogResult.OK)
                    //{
                        strFilePath = dialog.FileName;
                        strFileExt = Path.GetExtension(strFilePath);
                        if (strFileExt.CompareTo(".xls") == 0 || strFileExt.CompareTo(".xlsx") == 0)
                        {
                            try
                            {
                                dtResult = funReadExcel(strFilePath, strFileExt);
                                if (dtResult != null)
                                {
                                    for (int i = 0; i < dtResult.Rows.Count; i++)
                                    {
                                        if (dtResult.Rows[i]["ErackID"].ToString() != cboErackID.Text)
                                        {
                                            MessageBox.Show("ErackID[" + dtResult.Rows[i]["ErackID"].ToString() + "] is not SAME as [" + cboErackID.Text + "]","Error" ,MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                            return;
                                        }
                                        else
                                        {
                                            string strCmd = "UPDATE SHELF_INFO SET CycleID = '" + dtResult.Rows[i]["OccupyID"].ToString() + "' WHERE ShelfID = '" + dtResult.Rows[i]["ShelfID"].ToString() + "'";
                                            bool blnOK = ClsMSSQL.SqlNonQuery(strCmd, Global.g_INI.DB.ConnStr);
                                        }
                                    }

                                }
                                lblMSG.Text = "Import Success";
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                lblMSG.Text = "Import Fail";
                                return;
                            }
                        Create_InventoryResult();
                        }
                        else
                        {
                            MessageBox.Show("File Format is not Support!", "Hint", MessageBoxButtons.OK , MessageBoxIcon.Error);
                             return;
                        }
                    //}
                }
            }
          
        }

        private DataTable funReadExcel(string strFilePath, string strFileExt)
        {
            string strConn = string.Empty;
            DataTable dtTmp = new DataTable();
            if (strFileExt.CompareTo(".xls") == 0)
                strConn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source="+strFilePath+"; Extended Properties='Excel 8.0;HRD=YES;IMEX=1';";   //支援2007以下的版本
            else
                strConn = @"provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilePath + "; Extended Properties='Excel 12.0;HRD=NO';";   //支援2007以下的版本
            using (OleDbConnection con = new OleDbConnection(strConn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con);  //從sheet1 讀取資料
                    oleAdpt.Fill(dtTmp);
                }
                catch (Exception ex) 
                {
                    ClsLog.TraceLog(ex.ToString(), EnuLogType.Exception);
                }
            }
            return dtTmp;
        }

        private void txtItemID_KeyDown(object sender, KeyEventArgs e)
        {
            bool blnReOK = false;
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    if (!Global.g_INI.Common.PassHostTrx)
                    {
                        txtItemDetail.ReadOnly = true;
                        //加入W2106詢問物料詳細資訊
                        W2106_queryItemDetail_I jcsIn = new W2106_queryItemDetail_I();
                        W2106_queryItemDetail_O jcsOut = new W2106_queryItemDetail_O();
                        jcsIn.LINE = Global.g_INI.Common.LINE;
                        jcsIn.FLOOR = Global.g_INI.Common.FLOOR;
                        jcsIn.SYSTEM_ID = Global.g_INI.Common.SystemID;
                        jcsIn.CARR_ID = txtItemID.Text;

                        string strUrl = "http://localhost:14702/WebService11.asmx";     //WebService的http形式的地址 
                        string strNamespace = @"http://tempuri.org/";                   //欲呼叫的WebService的命名空間 
                        string strClassname = "LCSWCFService";                   //欲呼叫的WebService的類名（不包括命名空間前綴） 
                        string strMethodname = "queryItemDetail";                           //欲呼叫的WebService的方法名 
                        object[] objArgs = new object[1];                               //參數列表 

                        objArgs[0] = JsonConvert.SerializeObject(jcsIn, new JsonSerializerSettings() { StringEscapeHandling = StringEscapeHandling.EscapeNonAscii });
                        strUrl = Global.g_INI.Common.HOST_WebServiceUrl.Trim();
                        strNamespace = Global.g_INI.Common.HOST_WebServiceClassName;
                        strClassname = Global.g_INI.Common.HOST_WebServiceClassName;

                        object objReturnValue = ClsGlobalFunc.InvokeWebservice(strUrl, strNamespace, strClassname, strMethodname, objArgs);

                        if (objReturnValue != null)
                        {
                            string strReJson = objReturnValue.ToString().Trim();
                            if (!string.IsNullOrEmpty(strReJson))
                            {
                                blnReOK = true;
                                // 透過Json.NET反序列化為物件
                                try { jcsOut = JsonConvert.DeserializeObject<W2106_queryItemDetail_O>(strReJson); }
                                catch (Exception ex)
                                {
                                    blnReOK = false;
                                    ClsLog.TraceLog(jcsOut.GetType().Name + " 解碼失敗，strJson=[" + strReJson + "],ErrMsg=[" + ex.Message.ToString() + "]", EnuLogType.Error);
                                }

                                if (blnReOK)
                                {
                                    if (jcsOut.STATUS.Contains("OK"))
                                    {
                                        ClsLog.TraceLog(jcsOut.GetType().Name + " Reply OK，strJson=[" + strReJson + "]", EnuLogType.Trace);
                                        txtItemDetail.Text = jcsOut.ITEM_DETAIL;
                                    }
                                    else
                                    {
                                        ClsLog.TraceLog(jcsOut.GetType().Name + " Reply NG [" + jcsOut.MESSAGE.ToString() + "]，strJson=[" + strReJson + "]", EnuLogType.Error);
                                    }
                                }
                            }
                        }

                    }
                    else
                    {
                        //測試字串
                        MessageBox.Show("OFFLINE NOW!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        txtItemDetail.ReadOnly = false;
                    }
                }
                catch (Exception ex)
                {
                    ClsLog.TraceLog(ex.Message.ToString(), EnuLogType.Exception);
                    return;
                }
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            ERACK_EVENT_HISTORY erack_event_history = new ERACK_EVENT_HISTORY();
            //防呆1- ShelfID. EventType 不能為空
            if (string.IsNullOrEmpty(cboEventType.Text) || string.IsNullOrEmpty(txtShelfID.Text))
            {
                MessageBox.Show("EventType or ShelfID Must be Key in !(事件類型和儲位編號必填!)", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //防呆2- 根據EventType 檢查 ItemID 是否可以空值和檢查是否允許存取
            else
            {
                for (int i = 0; i < ucShelf.Length; i++)
                {
                    if (txtShelfID.Text == ucShelf[i].ShelfID)
                    {
                        if (cboEventType.Text == "LOAD" || cboEventType.Text == "CLEAR")
                        {
                            if (string.IsNullOrEmpty(txtItemID.Text))
                            {
                                MessageBox.Show("Must be key in ItemID!(必填物料編號!", "Hint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                            else
                            {
                                if (cboEventType.Text == "LOAD")
                                {
                                    if (ucShelf[i].Status == "EMPTY" || (ucShelf[i].Status.Contains("RESERVED") && ucShelf[i].ItemID == txtItemID.Text))
                                        funUPDShelfInfo();
                                    //else
                                    //    funErrorMsgShow();
                                }
                                else if (cboEventType.Text == "CLEAR")
                                {
                                    //if (ucShelf[i].Status == "FULL")
                                        funUPDShelfInfo();
                                    //else
                                    //    funErrorMsgShow();

                                }
                            }
                        }
                    }
                }
                funLoadingInventoryReuslt();
                funResetKeyInField();
            }
            // funLoadingInventoryReuslt();
         
        }

        private void funResetKeyInField()
        {
            cboEventType.SelectedIndex = -1;
            txtShelfID.Clear();
            txtItemID.Clear();
            txtItemDetail.Clear();
            txtMemo.Clear();
            lblCycleID.Text = "";
        }

        private void funUPDShelfInfo()
        {
            //EMPTY && Enable => LOAD
            string strCmd = string.Empty;
            bool blnOK = false;
            ERACK_EVENT_HISTORY erack_event_history = new ERACK_EVENT_HISTORY();

            try
            {
                string strStatus = string.Empty;
                string strManual = string.Empty;
                string strItemID = string.Empty;
                string strItemDetail = string.Empty;
                string strMemo = string.Empty;

                if (cboEventType.Text == "LOAD")
                {
                    strStatus = "FULL";
                    strItemID = txtItemID.Text;
                    strItemDetail = txtItemDetail.Text;
                    strMemo = txtMemo.Text;
                }
                else if (cboEventType.Text == "UNLOAD" || cboEventType.Text == "INITIAL")
                {
                    strStatus = "EMPTY";
                }

                strCmd = "UPDATE SHELF_INFO SET Status = '" + strStatus + "',Manual = 'Y',Match_Type='0', ";
                strCmd += "OccupyID = '" + strItemID + "',";
                strCmd += "ItemDetail='" + strItemDetail + "',";
                strCmd += "Memo = '" + strMemo + "',";
                strCmd += "UpdateTime= '" + ClsGlobalFunc.GetFormatDT() + "' ";
                strCmd += "WHERE ShelfID = '" + txtShelfID.Text + "'";

                blnOK = ClsMSSQL.SqlNonQuery(strCmd, Global.g_INI.DB.ConnStr);

                erack_event_history.LINE = Global.g_INI.Common.LINE;
                erack_event_history.FLOOR = Global.g_INI.Common.FLOOR;
                erack_event_history.SystemID = Global.g_INI.Common.SystemID;
                erack_event_history.ErackType = "ERACK";
                erack_event_history.ErackID = cboErackID.Text;
                erack_event_history.ShelfID = txtShelfID.Text;
                erack_event_history.FoupID = txtItemID.Text;
                erack_event_history.EventType = cboEventType.Text;
                erack_event_history.ACK = "N";
                erack_event_history.ErrorMsg = "Inventory Event:["+cboEventType.Text+"]";
                erack_event_history.CreateTime = ClsGlobalFunc.GetFormatDT();
                erack_event_history.ItemDetail = txtItemDetail.Text;
                erack_event_history.Memo = txtMemo.Text;


                if (blnOK && funInsertErackEventHistory(erack_event_history))
                {
                    lblMSG.Text = "[ShelfID]" + txtShelfID.Text + " [" + cboEventType.Text + "] COMPLETE!";
                }
            }
            catch (Exception ex)
            {
                ClsLog.TraceLog(ex.ToString(), EnuLogType.Exception);
            }



        }

        private void funErrorMsgShow()
        {
            MessageBox.Show("ShelfID[" + txtShelfID.Text + "] Not Allow to " + cboEventType.Text);
            return;
        }

        private bool funInsertErackEventHistory(ERACK_EVENT_HISTORY event_hisroy)
        {
            bool blnOK = false;
            try
            {
                string strCmd_I = "INSERT INTO [dbo].[ERACK_EVENT_HISTORY] ([LINE],[SystemID],[ErackType],[ErackID]" +
               ",[ShelfID],[CartID],[LocationID],[FoupID],[EventType],[ACK],[ErrorMsg],[CreateTime],[ItemDetail],[Memo])" +
               " VALUES ('" + event_hisroy.LINE + "','" + event_hisroy.SystemID + "','" + event_hisroy.ErackType + "','" + event_hisroy.ErackID + "','" + event_hisroy.ShelfID +
               "','" + event_hisroy.CartID + "','" + event_hisroy.LoationID + "','" + event_hisroy.FoupID + "','" + event_hisroy.EventType +
               "','" + event_hisroy.ACK + "','" + event_hisroy.ErrorMsg + "','" + event_hisroy.CreateTime + "','" + event_hisroy.ItemDetail +
               "','" + event_hisroy.Memo + "')";
                blnOK = ClsMSSQL.SqlNonQuery(strCmd_I, Global.g_INI.DB.ConnStr);
            }
            catch (Exception ex)
            {
                ClsLog.TraceLog(ex.ToString(), EnuLogType.Exception);
            }
            return blnOK;
        }

        private void btn_SaveChange_Click(object sender, EventArgs e)
        {
            string strCmd = string.Empty;
            bool blnOK = false;
            DataTable dtResult = new DataTable();
            
            strCmd = "SELECT * FROM SHELF_INFO WHERE ErackID= '"+cboErackID.Text+"'";
            blnOK = ClsMSSQL.GetDBData(strCmd, Global.g_INI.DB.ConnStr, ref dtResult);

            if (blnOK && dtResult.Rows.Count > 0)
            {
                MessageBox.Show("Please Input Inventory Result FileName, First !!", "Hint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                int intFileType = 2;
                string sFileName = objExp.funSaveFile("InventoryResult_" + DateTime.Today.ToString("yyyyMMdd") + "_" + cboErackID.Text, out intFileType);

                if (string.IsNullOrWhiteSpace(sFileName))
                {
                MessageBox.Show("Please Input the FileName First", "Hint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (ClsExport.funExpExcel(sFileName, dtResult))
                {
                    lblMSG.Text = "Export Success!! It will Unlock Erack";
                    funUpdateUnLock();
                }
                funLoadingShelfInfo(cboErackID.Text);
            }
               

        }

        private void funLoadingShelfInfo(string strErackID)
        {
            string strCmd = string.Empty;
            bool blnOK = false;
            DataTable dtTMP = new DataTable();
            try
            {
                strCmd = "SELECT * FROM SHELF_INFO WHERE ErackID ='" + strErackID + "'";
                blnOK = ClsMSSQL.GetDBData(strCmd, Global.g_INI.DB.ConnStr, ref dtTMP);
                if (blnOK && dtTMP.Rows.Count > 0)
                {
                    if (ucShelf != null)
                    {
                        this.panel_ShelfInfo.Controls.Clear();
                        ucShelf = null;
                    }

                    //載入格位排版
                    ucShelf = new ucShelf[12];
                    for (int i = 0; i < ucShelf.Length; i++)
                    {
                        ucShelf[i] = new ucShelf();
                        panel_ShelfInfo.Controls.Add(ucShelf[i]);
                        ucShelf[i].Width = (this.Width - 450) / 4;
                        ucShelf[i].Height = (int)(ucShelf[i].Width * 0.9);
                        ucShelf[i].Top = 0 + (i / 4) * ucShelf[i].Height;
                        ucShelf[i].Left = 3 + (i % 4) * ucShelf[i].Width;
                        ucShelf[i].Visible = false;
                        ucShelf[i].Tag = i;
                        ucShelf[i].MouseClick += new System.Windows.Forms.MouseEventHandler(ucShelf_MouseClick);

                    }
                    int intCol = dtTMP.Rows.Count / 3;
                    //載入格位資訊
                    for (int j = 0; j < dtTMP.Rows.Count; j++)
                    {
                        int intIdx = (j % intCol) + (j / intCol) * 4;
                        ucShelf[intIdx].ShelfID = dtTMP.Rows[j]["ShelfID"].ToString().Trim();
                        ucShelf[intIdx].Visible = true;

                        string strStatus = string.IsNullOrEmpty(dtTMP.Rows[j]["Status"].ToString()) ? string.Empty : dtTMP.Rows[j]["Status"].ToString().Trim();
                        bool blnDisable = false;  //是否為禁用儲位
                        if (!string.IsNullOrEmpty(dtTMP.Rows[j]["Enable"].ToString()))
                        {
                            if (dtTMP.Rows[j]["Enable"].ToString() == "Y")
                            {
                                blnDisable = false; //不是禁用儲位

                            }
                            else
                            {
                                blnDisable = true;  //是禁用儲位

                            }
                        }
                        string strPurpose = string.IsNullOrEmpty(dtTMP.Rows[j]["Purpose"].ToString()) ? string.Empty : dtTMP.Rows[j]["Purpose"].ToString().Trim();
                        string strCarrID = string.IsNullOrEmpty(dtTMP.Rows[j]["OccupyID"].ToString()) ? string.Empty : dtTMP.Rows[j]["OccupyID"].ToString().Trim();
                        string strReserved_CarrID = string.IsNullOrEmpty(dtTMP.Rows[j]["ReservedID"].ToString()) ? string.Empty : dtTMP.Rows[j]["ReservedID"].ToString().Trim();
                        string strManual = string.IsNullOrEmpty(dtTMP.Rows[j]["Manual"].ToString()) ? string.Empty : dtTMP.Rows[j]["Manual"].ToString().Trim();

                        //儲位狀態顏色設定
                        //if (strPurpose == "N")
                        //    ucShelf[intIdx].PurposeChange(strPurpose);
                        if (blnDisable)
                            ucShelf[intIdx].ShelfDisable(strCarrID);
                        else if (strStatus == "ALARM")
                            ucShelf[intIdx].ShelfAlarm(strStatus);
                        else
                        {
                            ucShelf[intIdx].ShelfUnLoad(strCarrID);
                            if (strStatus == "FULL")
                                ucShelf[intIdx].ShelfLoad(strCarrID);
                            else if (strStatus.Contains("RESERVE"))
                                ucShelf[intIdx].ShelfReserved(strReserved_CarrID);
                        }
                        //存取判斷上下料會用到的資料
                        ucShelf[intIdx].Status = strStatus;
                        ucShelf[intIdx].ItemID = strCarrID;
                    }
                }
            }
            catch (Exception ex)
            {
                ClsLog.TraceLog(ex.ToString(), EnuLogType.Exception);
                return;
            }
            finally
            {
                if (dtTMP != null)
                    dtTMP.Dispose();
            }
        }

        private void cboErackID_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (rbt_Manual.Checked)
            //{
                if (!string.IsNullOrEmpty(cboErackID.Text))
                    funLoadingShelfInfo(cboErackID.Text);
            //}
            funResetKeyInField();
        }

        private void btnLock_Click(object sender, EventArgs e)
        {
            if (funGetErackStatus() == "DISABLE")
            {
                DialogResult CheckLock = MessageBox.Show("It will UNLOCK the Erack! Continue?", "Hint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                if (CheckLock == DialogResult.OK)
                    funUpdateUnLock();
                else
                    return;
            }
            else if (funGetErackStatus() == "ENABLE")
            {
                DialogResult CheckUnLock = MessageBox.Show("It will LOCK the Erack! Continue?", "Hint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                if (CheckUnLock == DialogResult.OK)
                    funUpdateLock();
                else
                    return;
            }
        }

        private string funGetErackStatus()
        {
            string strCmd = string.Empty;
            bool blnOK = false;
            DataTable dtTmp = new DataTable();
            string strRtnStatus = string.Empty;
            strCmd = "SELECT * FROM ERACK_INFO WHERE ErackID = '"+cboErackID.Text+"'";
            try
            {
                blnOK = ClsMSSQL.GetDBData(strCmd, Global.g_INI.DB.ConnStr, ref dtTmp);
                if (blnOK)
                {
                    if (dtTmp.Rows.Count > 0)
                    {
                        strRtnStatus = dtTmp.Rows[0]["Status"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                ClsLog.TraceLog(ex.Message, EnuLogType.Exception);
            } 
            finally
            {
                if (dtTmp != null) 
                    dtTmp.Dispose();
            }
            return strRtnStatus;

            
        }
    }
}
