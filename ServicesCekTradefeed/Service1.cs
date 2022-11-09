using Microsoft.Build.Tasks;
using Microsoft.Reporting.WebForms;
using Newtonsoft.Json;
using OpenQA.Selenium;
using RestSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Types.Enums;
using Warning = Microsoft.Reporting.WebForms.Warning;
using OpenQA.Selenium.Chrome;
using System.Data.OleDb;
using OpenQA.Selenium.Interactions;
using System.Text.RegularExpressions;
using IronPdf;
using IronOcr;
using System.Reflection;
using Telegram.Bot.Types;
using File = System.IO.File;
using System.Security.Cryptography;

namespace ServicesCekTradefeed
{
    public partial class Service1 : ServiceBase
    {
        public static List<irca_parameter_new> new_data = new List<irca_parameter_new>();
        private static TelegramBotClient bot = new TelegramBotClient("2144239635:AAFjcfn_GdHP4OkzzZomaZt4XbwpHDGyR-U");
        public static string connectionString = "Data Source=KBIDC-SKD-DBMS.ptkbi.com;Initial Catalog=SKD;Persist Security Info=True;User ID=saskd;Password=P@ssw0rd123";

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            WriteToFile($"Start listening!!");
            bot.StartReceiving(Array.Empty<UpdateType>());

            System.Timers.Timer timer = new System.Timers.Timer();

            //timer.Interval = 1000;

            bot.OnUpdate += BotOnUpdateReceived;

            //timer.Start();
        }
        public void OnTimer(object sender, ElapsedEventArgs args)
        {
        }
        private static async void BotOnUpdateReceived(object sender, UpdateEventArgs e)
        {
            try
            {
                
                //GET BUSINES DATE
                DateTime buseinesdate = DateTime.Now;
                var querry = "SELECT DateValue FROM SKD.Parameter WHERE Code = 'BusinessDate'";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlCommand command = new SqlCommand(querry, connection);
                    connection.Open();
                    command.CommandTimeout = 1800;
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        buseinesdate = Convert.ToDateTime((reader["DateValue"]));
                    }
                    reader.Close();
                    connection.Dispose();
                    connection.Close();
                }
                //END

                var message = e.Update.Message;
                WriteToFile(message.Type.ToString());
                if (message.ReplyToMessage != null)
                {
                    //perpanjang contract
                    if (message.ReplyToMessage.Text.Contains("/e Commodity Code"))
                    {
                        insertContract(message.Chat.Id, message.ReplyToMessage.Text, message.Text);
                        monitoringServices("DOP_EODTele", "I", "Service untuk insert contract SKD");

                    }
                    //get report monthly
                    if (message.ReplyToMessage.Text.Contains("Report Monthly SKD"))
                    {
                        generetaRptCluster(message.Chat.Id, message.Text);
                        monitoringServices("DOP_EODTele", "I", "Service untuk get report clearing member SKD");
                    }
                    //get fee paln
                    if (message.ReplyToMessage.Text.Contains("fee paln"))
                    {
                        gettotalpaln(message.Chat.Id, message.Text);
                        monitoringServices("DOP_EODTele", "I", "Service untuk get data total PALN");
                    }
                }
                if (message.Type == MessageType.Text)
                {
                    var text = message.Text;
                    WriteToFile(text);
                    if (text == "/cektradefeed@Automation_KBI_Bot")
                    {
                        cektradefeed(message.Chat.Id, buseinesdate);
                        monitoringServices("DOP_EODTele", "I", "Service untuk cek tradefeed SKD");

                    }
                    if (text == "/autoapproveexceptions@Automation_KBI_Bot")
                    {
                        cekSP(message.Chat.Id, buseinesdate);
                        monitoringServices("DOP_EODTele", "I", "Service untuk auto approve exceptions SKD");

                    }
                    if (text == "/eodskd@Automation_KBI_Bot")
                    {
                        eodPBK(message.Chat.Id, buseinesdate);
                        monitoringServices("DOP_EODTele", "I", "Service untuk EOD SKD");

                    }
                    if (text == "/sodskd@Automation_KBI_Bot")
                    {
                        sodPBK(message.Chat.Id, buseinesdate);
                        monitoringServices("DOP_EODTele", "I", "Service untuk SOD SKD");

                    }
                    if (text == "/getreportdireksi@Automation_KBI_Bot")
                    {
                        getreportdireksi(message.Chat.Id, buseinesdate);
                        monitoringServices("DOP_EODTele", "I", "Service untuk get report direksi SKD");

                    }
                    if (text == "/publishreportskd@Automation_KBI_Bot")
                    {
                        publishReport(message.Chat.Id);
                        monitoringServices("DOP_EODTele", "I", "Service untuk publish report SKD");

                    }
                    if (text == "/replicatereportskd@Automation_KBI_Bot")
                    {
                        replicatedReport(message.Chat.Id);
                        monitoringServices("DOP_EODTele", "I", "Service untuk replicated report SKD");

                    }
                    if (text == "/sendemailskd@Automation_KBI_Bot")
                    {
                        sendReport(message.Chat.Id);
                        monitoringServices("DOP_EODTele", "I", "Service untuk send report SKD");

                    }
                    if (text == "/getdatapaln@Automation_KBI_Bot")
                    {
                        getdataPaln(message.Chat.Id);
                        monitoringServices("DOP_EODTele", "I", "Service untuk get data PALN");

                    }
                    if (text == "/rptshortagedaily@Automation_KBI_Bot")
                    {
                        rptShortageDailyDKA(message.Chat.Id);
                        monitoringServices("DOP_EODTele", "I", "Service untuk generate report shortage daily");

                    }
                    if (text == "/pengecekenreport@Automation_KBI_Bot")
                    {
                        pengecekanReportTerkirim(message.Chat.Id);
                        monitoringServices("DOP_EODTele", "I", "Service untuk pengecekan report SKD");

                    }
                    if (text == "/getreportmonthly@Automation_KBI_Bot")
                    {
                        bot.SendTextMessageAsync(message.Chat.Id, "*Report Monthly SKD*\nPlease answer this message\n\n*januari\n*februari\n*maret\n*april\n*mei\n*juni\n*juli\n*agustus\n*september\n*oktober\n*november\n*desember", ParseMode.Markdown);
                        monitoringServices("DOP_EODTele", "I", "Service untuk request report clearing member monthly");
                    }
                    if (text == "/insertdatapaln@Automation_KBI_Bot")
                    {
                        insert_data_paln(message.Chat.Id);
                        monitoringServices("DOP_EODTele", "I", "Service untuk request report clearing member monthly");
                    }
                    if (text == "/cektotalfeepaln@Automation_KBI_Bot")
                    {
                        bot.SendTextMessageAsync(message.Chat.Id, "*fee paln*\nPlease answer this message\n\n*startdate (yyyy-MM-dd)\n*endDate (yyyy-MM-dd)", ParseMode.Markdown);
                        monitoringServices("DOP_EODTele", "I", "Service untuk request check total fee PALN");
                    }
                }
                    //jika mengirim dokumen
                    if (message.Type == MessageType.Document)
                    {
                        WriteToFile("dokumen masuk kesini");
                        var file = await bot.GetFileAsync(message.Document.FileId);
                        FileStream fs = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "\\" + message.Document.FileName, FileMode.Create);
                        await bot.DownloadFileAsync(file.FilePath, fs);
                        fs.Close();
                        fs.Dispose();
                        WriteToFile(fs.Name);
                        if (message.Document.FileName.ToLower().Contains("beda bunga curr"))
                        {
                            insertIrca(message.Chat.Id, fs.Name);
                            monitoringServices("DOP_EODTele", "I", "Service untuk insert irca SKD");

                        }
                        if (message.Document.FileName.Contains("S287O"))
                        {
                            insert_paln_sfi_1(message.Chat.Id, fs.Name);
                            monitoringServices("DOP_EODTele", "I", "Service untuk insert PALN SFI");
                        }
                        if (message.Document.FileName.Contains("PT_Straits_Futures_Indonesia"))
                        {
                            insert_paln_sfi_2(message.Chat.Id, fs.Name);
                            monitoringServices("DOP_EODTele", "I", "Service untuk insert PALN SFI");
                        }
                        if (message.Document.FileName.Contains("S289O9000A"))
                        {
                            insert_paln_pg_1(message.Chat.Id, fs.Name);
                            monitoringServices("DOP_EODTele", "I", "Service untuk insert PALN PG");
                        }
                        if (message.Document.FileName.Contains("ALPACA"))
                        {
                            insert_paln_pg_2(message.Chat.Id, fs.Name);
                            monitoringServices("DOP_EODTele", "I", "Service untuk insert PALN PG");
                        }
                        if (message.Document.FileName.Contains("Berjangka_Daily"))
                        {
                            insert_paln_pg_3(message.Chat.Id, fs.Name);
                            monitoringServices("DOP_EODTele", "I", "Service untuk insert PALN PG");
                        }
                    }
            }
            catch (Exception ex)
            {
                WriteToFile("prosess fail " + ex.Message + " " + DateTime.Now.ToString());
            }
        }
        public static void cektradefeed(long chat_id, DateTime businessDate)
        {
            var dr_tr = new DataSet1TableAdapters.DataTable1TableAdapter();
            var dt_primer = dr_tr.GetDataBBJPRIMER(businessDate);
            var dt_spa = dr_tr.GetDataByBBJSPA(businessDate);
            bot.SendTextMessageAsync(chat_id, "Tanggal: " + businessDate.ToString("dd MMM yyyy") + "\n\nBBJ PRIMER \nTotal Trade: " + dt_primer[0].totaltradefeed + "\nTotal Lot: " + dt_primer[0].lot + "\n\nBBJ SPA\nTotal Trade: " + dt_spa[0].totaltradefeed + "\nTotal Lot: " + dt_spa[0].lot, ParseMode.Markdown);
        }
        public static void cekSP(long chat_id, DateTime businessDate)
        {
            var dr_validasi = new DataSet1TableAdapters.Validation_SettlementPrice_ExceptionsTableAdapter();
            var dt_validasi = dr_validasi.GetData(businessDate);
            WriteToFile("Validasi EOD " + DateTime.Now.ToString("hh:mm:ss"));

            var x = 1;
            var dr_trade_exception = new DataSet1TableAdapters.TradefeedExceptionTableAdapter();
            var dr_approveAll = new DataSet1TableAdapters.QueriesTableAdapter();
            //var dt_trade_excption = dr_trade_exception.GetDataByStatus(DateTime.Now.AddDays(-1));
            try
            {
                string printOutput = "";
                string connectionStringSKDSP = "Data Source=KBIDC-SKD-DBMS.ptkbi.com;Initial Catalog=SKDSP;Persist Security Info=True;User ID=saskd;Password=P@ssw0rd123";
                string queryString = "EXEC [dbo].[uspMigrateSettlementPriceRedo] '" + businessDate.ToString("yyyy-MM-dd") + "'";

                using (SqlConnection connection = new SqlConnection(
    connectionStringSKDSP))
                {
                    SqlCommand command = new SqlCommand(queryString, connection);
                    connection.Open();
                    command.CommandTimeout = 1800;
                    connection.InfoMessage += (object obj, SqlInfoMessageEventArgs e) =>
                    {
                        printOutput += e.Message;
                    };
                    SqlDataReader reader = command.ExecuteReader();

                    connection.Dispose();
                    connection.Close();
                }
                dr_approveAll.uspTradeFeedExceptionApproveAll(2, businessDate, "A", "oke", "Robot KBI");
                bot.SendTextMessageAsync(chat_id, "Business Date : " + businessDate.ToShortDateString() + "\n\nSettlement Price " + dt_validasi[0].count_sp + " Record\nTradefeed Exceptions " + dt_validasi[0].count_ex + " Record\n\nSuccessfully run the auto approve command\nSuccess execute [uspMigrateSettlementPriceRedo]\n\nTime Stamp :" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Proses auto approve fail \n\n" + ex.Message + " \n\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
            }

        }
        public static void eodPBK(long chat_id, DateTime businessDate)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Proses EOD start " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

                string queryStringEOD = "EXEC [SKD].[EOD_ProcessEOD] '" + businessDate.ToString("yyyy-MM-dd") + "', 'EOD','GLOBAL',null,null,null,'Robot KBI','10.15.2.60'";
                SendTelegram("-1001641017183", "[" + DateTime.Now.ToString("dd MMM yyyy") + "] [" + DateTime.Now.ToString("hh:mm:ss") + "]\nLog : Processing EOD On Progress. . .\nQuery : " + queryStringEOD);
                WriteToFile(queryStringEOD);
                using (SqlConnection connection = new SqlConnection(
                           connectionString))
                {
                    SqlCommand command = new SqlCommand(queryStringEOD, connection);
                    connection.Open();
                    command.CommandTimeout = 3600000;
                    SqlDataReader reader = command.ExecuteReader();

                    connection.Close();
                }
                bot.SendTextMessageAsync(chat_id, "Proses EOD succes " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                WriteToFile("EOD Success " + DateTime.Now);

            }
            catch (Exception x)
            {
                bot.SendTextMessageAsync(chat_id, "Proses EOD fail \n\n" + x.Message + " \n\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
            }

        }
        public static void sodPBK(long chat_id, DateTime businessDateparam)
        {
            List<data_shortage_direksi> data_direksi = new List<data_shortage_direksi>();
            try
            {

                SendTelegram("-1001641017183", "[" + DateTime.Now.ToString("dd MMM yyyy") + "] [" + DateTime.Now.ToString("hh:mm:ss") + "]\nLog : Processing SOD Start");
                bot.SendTextMessageAsync(chat_id, "Proses Start Of Day start " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

                DateTime dateValue = DateTime.Now.Date;

                for (int i = 0; i < 6; i++)
                {
                    if (IsHoliday(dateValue))
                    {
                        string queryStringEOD = "EXEC [SKD].[EOD_ProcessEOD] '" + dateValue.ToString("yyyy-MM-dd") + "', 'EOD','GLOBAL',null,null,null,'Robot KBI','10.10.10.231'";

                        SendTelegram("-1001641017183", "[" + DateTime.Now.ToString("dd MMM yyyy") + "] [" + DateTime.Now.ToString("hh:mm:ss") + "]\nLog : Processing SOD On Progress\nQuery : " + queryStringEOD);

                        WriteToFile("SOD : " + queryStringEOD);
                        using (SqlConnection connection = new SqlConnection(
                                   connectionString))
                        {
                            SqlCommand command = new SqlCommand(queryStringEOD, connection);
                            connection.Open();
                            command.CommandTimeout = 1800;
                            SqlDataReader reader = command.ExecuteReader();
                            connection.Dispose();
                            connection.Close();
                        }

                        dateValue = dateValue.AddDays(1);
                    }
                    else
                    {
                        break;
                    }
                }

                bot.SendTextMessageAsync(chat_id, "SOD Date Value : " + dateValue.ToString("dd MMM yyyy"), ParseMode.Markdown);

                WriteToFile("SOD Date " + dateValue + " - Time Stamp : " + DateTime.Now.ToString("hh:mm:ss"));
                var dr_sod = new DataSet1TableAdapters.ParameterTableAdapter();
                var dt_sod = dr_sod.GetDataById();
                dt_sod[0].DateValue = dateValue;
                dt_sod[0].LastUpdatedBy = "Robot KBI";
                dt_sod[0].LastUpdatedDate = DateTime.Now;
                dr_sod.Update(dt_sod);

                var dt_lastEOD = dr_sod.GetDataByLastEOD();
                dt_lastEOD[0].DateValue = DateTime.Now.Date.AddDays(-1);
                dt_lastEOD[0].LastUpdatedBy = "Robot KBI";
                dt_lastEOD[0].LastUpdatedDate = DateTime.Now;
                dr_sod.Update(dt_lastEOD);

                bot.SendTextMessageAsync(chat_id, "Prosess Start Of Day Success " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

                WriteToFile("Sukses SOD");

            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Proses SOD PBK fail \n\n" + ex.Message + " \n\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                monitoringServices("DOP_EODTele", "E", "Proses SOD PBK fail " + ex.Message);

            }
            try
            {
                List<string> summary = new List<string>();

                SendTelegram("-1001641017183", "[" + DateTime.Now.ToString("dd MMM yyyy") + "] [" + DateTime.Now.ToString("hh:mm:ss") + "]\nLog : Processing Restrictions Shortage Start");
                bot.SendTextMessageAsync(chat_id, "Process Pembatasan Shortage Start " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

                var dr_srtge = new DataSet1TableAdapters.uspRptDailyShortageSummaryTableAdapter();
                var dt_srtge = dr_srtge.GetData(DateTime.Now.Date);

                //var dr_srtge = new DataSetDevTableAdapters.uspRptDailyShortageSummaryTableAdapter();
                //var dt_srtge = dr_srtge.GetData(Convert.ToDateTime("2010-08-10"));

                int i = 1;
                int z = 0;
                foreach (var item in dt_srtge)
                {
                    summary.Add(i + ". Nama : " + item.Name + ", Total Equivalent : (" + (item.TotalEkuivalen * (-1)).ToString("#,##0.####") + ") , Status : " + item.StatusSuspensi);
                    if (item.TotalEkuivalen < -500000000 && item.StatusSuspensi == "Active")
                    {
                        var data = new data_shortage_direksi
                        {
                            nama = item.Name,
                            nominal = item.TotalEkuivalen
                        };
                        data_direksi.Add(data);
                        z++;
                        DataSet1 dsClearingMemberStatus = new DataSet1();
                        var dr_cm_status = new DataSet1TableAdapters.ClearingMemberSuspenseStatusTableAdapter();
                        var dr_suspend = new DataSet1TableAdapters.ClearingHouseSuspenseTableAdapter();
                        var dr_cm = new DataSet1TableAdapters.ClearingMemberTableAdapter();
                        var dr_param = new DataSet1TableAdapters.ParameterTableAdapter();


                        //DataSetDev dsClearingMemberStatus = new DataSetDev();
                        //var dr_cm_status = new DataSetDevTableAdapters.ClearingMemberSuspenseStatusTableAdapter();
                        //var dr_suspend = new DataSetDevTableAdapters.ClearingHouseSuspenseTableAdapter();
                        //var dr_cm = new DataSetDevTableAdapters.ClearingMemberTableAdapter();
                        //var dr_param = new DataSet1TableAdapters.ParameterTableAdapter();


                        var maxSuspendNo = 0;
                        var businessDate = dr_param.GetDataByBusinessDate()[0].DateValue;
                        //var businessDate = Convert.ToDateTime("2010-08-10"); // dr_param.GetDataByBusinessDate()[0].DateValue;

                        //====For Exchange Primer====
                        var exchangeId = 1;
                        dsClearingMemberStatus.ClearingHouseSuspense.Clear();
                        //var dt_suspend = dr_suspend.GetDataByMaxSuspendNo(exchangeId, Convert.ToDateTime("2010-08-10"));
                        maxSuspendNo = dr_suspend.FillByMaxSuspendNo(dsClearingMemberStatus.ClearingHouseSuspense, exchangeId, businessDate);
                        if (dsClearingMemberStatus.ClearingHouseSuspense.Rows.Count > 0)
                        {
                            //maxSuspendNo = ((DataSetDev.ClearingHouseSuspenseRow)dsClearingMemberStatus.ClearingHouseSuspense.Rows[0]).SuspenseNo;

                            maxSuspendNo = ((DataSet1.ClearingHouseSuspenseRow)dsClearingMemberStatus.ClearingHouseSuspense.Rows[0]).SuspenseNo;
                        }
                        else
                        {
                            maxSuspendNo = 0;
                        }

                        var suspendTime = DateTime.Now;
                        var cmCode = item.Code;
                        var cmId = dr_cm.GetDataByCode(item.Code)[0].CMID;
                        var suspendType = "L";
                        var description = "Pembatasan pagi";

                        maxSuspendNo = maxSuspendNo + 1;
                        var newId = -1;
                        Object retObj = null;

                        retObj = dr_suspend.InsertQuery(businessDate, exchangeId, maxSuspendNo, suspendTime
                                 , null, cmId, suspendType, description);
                        newId = Convert.ToInt32(retObj);
                        dr_cm_status.Insert(businessDate, exchangeId, cmId, "C", newId, suspendType, suspendTime, "N", "Y");

                        //====For Exchange SPA======
                        exchangeId = 2;
                        maxSuspendNo = 0;
                        dsClearingMemberStatus.ClearingHouseSuspense.Clear();
                        //dt_suspend = dr_suspend.GetDataByMaxSuspendNo(exchangeId, Convert.ToDateTime("2010-08-10"));
                        maxSuspendNo = dr_suspend.FillByMaxSuspendNo(dsClearingMemberStatus.ClearingHouseSuspense, exchangeId, businessDate);
                        if (dsClearingMemberStatus.ClearingHouseSuspense.Rows.Count > 0)
                        {
                            //maxSuspendNo = ((DataSetDev.ClearingHouseSuspenseRow)dsClearingMemberStatus.ClearingHouseSuspense.Rows[0]).SuspenseNo;
                            maxSuspendNo = ((DataSet1.ClearingHouseSuspenseRow)dsClearingMemberStatus.ClearingHouseSuspense.Rows[0]).SuspenseNo;
                        }
                        else
                        {
                            maxSuspendNo = 0;
                        }

                        maxSuspendNo = maxSuspendNo + 1;

                        retObj = dr_suspend.InsertQuery(businessDate, exchangeId, maxSuspendNo, suspendTime
                                 , null, cmId, suspendType, description);
                        newId = Convert.ToInt32(retObj);
                        dr_cm_status.Insert(businessDate, exchangeId, cmId, "C", newId, suspendType, suspendTime, "N", "Y");
                    }
                    i++;
                }
                bot.SendTextMessageAsync(chat_id, String.Join("\n", summary), ParseMode.Markdown);
                bot.SendTextMessageAsync(chat_id, z + " Data berhasil di limit ", ParseMode.Markdown);

                bot.SendTextMessageAsync(chat_id, "Process Pembatasan Shortage Success " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

                WriteToFile(String.Join("\n", summary));

            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Proses SOD PBK fail \n\n" + ex.Message + " \n\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                monitoringServices("DOP_EODTele", "E", "Proses SOD PBK fail " + ex.Message);

            }
            try
            {
                bot.SendTextMessageAsync(chat_id, "Process Exec UspGenerateFasDFS Start " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

                var dr_dfs = new DataSet1TableAdapters.QueriesTableAdapter();
                dr_dfs.uspGenerateFasDFS(DateTime.Now.Date.AddDays(-1));

                bot.SendTextMessageAsync(chat_id, "Process Exec UspGenerateFasDFS Success " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

                WriteToFile("Process Exec UspGenerateFasDFS Sukses " + DateTime.Now.Date.AddDays(-1) + " - TimeStamp : " + DateTime.Now.ToString("hh:mm:ss"));
            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Proses UspGenerateFasDFS fail \n\n" + ex.Message + " \n\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                monitoringServices("DOP_EODTele", "E", "Proses SOD PBK fail " + ex.Message);

            }
            List<string> data_dir = new List<string>();
            int no = 1;
            foreach (var item in data_direksi)
            {
                data_dir.Add(no + ". Nama : " + item.nama + ", Total Equivalent : (" + (item.nominal * (-1)).ToString("#,##0.####") + ")");
                no++;
            }
            try
            {
                decimal nominal = 0;

                string queryStringRptDir = "EXEC [SKD].[uspRptDirectorTransactionSummary] '" + DateTime.Now.Date.AddDays(-1).ToString("yyyy-MM-dd") + "','" + DateTime.Now.Date.AddDays(-1).ToString("yyyy-MM-dd") + "'";

                using (SqlConnection connection = new SqlConnection(
               connectionString))
                {
                    SqlCommand command = new SqlCommand(queryStringRptDir, connection);
                    connection.Open();
                    command.CommandTimeout = 1800;
                    SqlDataReader reader = command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            nominal = Convert.ToDecimal(reader["AVERAGEINCLUDESS"].ToString());
                        }
                    }

                    connection.Dispose();
                    connection.Close();
                }
                bot.SendTextMessageAsync(chat_id, "Proses get report direksi start" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

                if (data_direksi.Count == 0)
                {
                    bot.SendTextMessageAsync(chat_id, "Yth. Bapak Direksi\nYth.Bapak / Ibu Koordinator\n\nBerikut laporan transaksi  Tanggal " + DateTime.Now.AddDays(-1).ToString("dd MMMM yyyy") + ":\n\nRata - rata kumulatif: " + nominal.ToString("#,##0.##") + "\nPembatasan: - \nTim DOP", ParseMode.Markdown);
                }
                else
                {
                    bot.SendTextMessageAsync(chat_id, "Yth. Bapak Direksi\nYth.Bapak / Ibu Koordinator\n\nBerikut laporan transaksi  Tanggal " + DateTime.Now.AddDays(-1).ToString("dd MMMM yyyy") + ":\n\nRata - rata kumulatif: " + nominal.ToString("#,##0.##") + "\nPembatasan: " + String.Join("\n", data_dir) + "\nTim DOP", ParseMode.Markdown);
                }
                var path = getRptImage("/SKDReport/RptDirectorAndFeeForWhatsapp");
                sendFileTelegram("-1001671146559", path);
            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Proses generate report direksi fail \n\n" + ex.Message + " \n\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                monitoringServices("DOP_EODTele", "E", "Proses SOD PBK fail " + ex.Message);

            }
        }
        public static void getreportdireksi(long chat_id, DateTime businessDateparam)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Process Generate Get Report Direksi " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                var path = getRptImage("/SKDReport/RptDirectorAndFeeForWhatsapp");
                sendFileTelegram("-1001671146559", path);
            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Process Generate Get Report Direksi Fail " + ex.Message + " " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

            }

        }
        public static void publishReport(long chat_id)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Proses generate report start : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

                IWebDriver ChromeDriver = new ChromeDriver();
                ChromeDriver.Navigate().GoToUrl("http://10.10.10.79/login.aspx?ReturnUrl=%2fdefault.aspx");

                IWebElement username = ChromeDriver.FindElement(By.Id("uiAuthLogin_UserName"));
                IWebElement password = ChromeDriver.FindElement(By.Id("uiAuthLogin_Password"));
                IWebElement submit_login = ChromeDriver.FindElement(By.Id("uiAuthLogin_LoginImageButton"));

                username.SendKeys("Hasto");
                password.SendKeys("Jakarta2019");
                submit_login.Click();
                WriteToFile("Sukses connect");
                ChromeDriver.Navigate().GoToUrl("http://10.10.10.79/ClearingAndSettlement/PublishReport2.aspx");

                IWebElement Date = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_CtlCalendarPickUpEOD_uiTxtCalendar"));

                IWebElement submit_publish = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_uiBtnPublish"));

                IJavaScriptExecutor js = (IJavaScriptExecutor)ChromeDriver;
                js.ExecuteScript("arguments[0].removeAttribute('readonly','readonly')", Date);

                WriteToFile("1");

                Date.Clear();
                var tgl_eod = DateTime.Now.AddDays(-1).ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture);
                //var tgl_eod = "01-Dec-2010";
                WriteToFile("2");
                Date.SendKeys(tgl_eod);
                WriteToFile("3");

                IWebElement replicate_report = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_cbReplicate"));
                IWebElement send_report = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_uiBtnPublish"));
                IWebElement loading = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_UpdateProgress1"));

                //Generate Report
                WriteToFile("4");
                replicate_report.Click();
                send_report.Click();
                WriteToFile("5");
                submit_publish.Submit();
                WriteToFile(submit_publish.GetAttribute("value"));
                WriteToFile(tgl_eod);
                Thread.Sleep(720000);
                if (loading.Displayed)
                {
                    bot.SendTextMessageAsync(chat_id, "Masih proses, sabar ya.. : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    bot.SendTextMessageAsync(chat_id, "Masih proses, sabar ya.. : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }

                ChromeDriver.Quit();

                WriteToFile("Berhasil generate publish report : " + DateTime.Now);
                bot.SendTextMessageAsync(chat_id, "Proses generate report selesai : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "gagal publish report: " + ex.Message + "\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
            }
        }
        public static void replicatedReport(long chat_id)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Proses replicated report start : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

                IWebDriver ChromeDriver = new ChromeDriver();
                ChromeDriver.Navigate().GoToUrl("http://10.10.10.79/login.aspx?ReturnUrl=%2fdefault.aspx");

                IWebElement username = ChromeDriver.FindElement(By.Id("uiAuthLogin_UserName"));
                IWebElement password = ChromeDriver.FindElement(By.Id("uiAuthLogin_Password"));
                IWebElement submit_login = ChromeDriver.FindElement(By.Id("uiAuthLogin_LoginImageButton"));

                username.SendKeys("Hasto");
                password.SendKeys("Jakarta2019");
                submit_login.Click();
                WriteToFile("Login sukses");
                ChromeDriver.Navigate().GoToUrl("http://10.10.10.79/ClearingAndSettlement/PublishReport2.aspx");

                IWebElement Date = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_CtlCalendarPickUpEOD_uiTxtCalendar"));

                IWebElement submit_publish = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_uiBtnPublish"));
                WriteToFile("1");
                IJavaScriptExecutor js = (IJavaScriptExecutor)ChromeDriver;
                WriteToFile("2");
                js.ExecuteScript("arguments[0].removeAttribute('readonly','readonly')", Date);
                WriteToFile("3");
                Date.Clear();
                var tgl_eod = DateTime.Now.AddDays(-1).ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture);
                Date.SendKeys(tgl_eod);
                IWebElement generate_report = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_cbGenerateReport"));
                IWebElement replicate_report = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_cbReplicate"));
                IWebElement send_report = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_cbPublish"));
                IWebElement loading = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_UpdateProgress1"));

                //Generate Report
                WriteToFile("4");
                generate_report.Click();
                send_report.Click();
                WriteToFile("5");
                submit_publish.Submit();

                Thread.Sleep(60000);
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);

                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);

                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);

                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);

                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);

                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);

                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);

                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);

                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);

                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);

                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);

                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);

                }

                ChromeDriver.Quit();

                WriteToFile("Berhasil replicated publish report : " + DateTime.Now);
                bot.SendTextMessageAsync(chat_id, "Proses replicated report selesai : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "gagal replicated publish report: " + ex.Message + "\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
            }

        }
        public static void sendReport(long chat_id)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Proses send report start : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                IWebDriver ChromeDriver = new ChromeDriver();
                ChromeDriver.Navigate().GoToUrl("http://10.10.10.79/login.aspx?ReturnUrl=%2fdefault.aspx");

                IWebElement username = ChromeDriver.FindElement(By.Id("uiAuthLogin_UserName"));
                IWebElement password = ChromeDriver.FindElement(By.Id("uiAuthLogin_Password"));
                IWebElement submit_login = ChromeDriver.FindElement(By.Id("uiAuthLogin_LoginImageButton"));

                username.SendKeys("Hasto");
                password.SendKeys("Jakarta2019");
                submit_login.Click();
                WriteToFile("Sukses Login 6281310215750 - 1548729946@g.us");

                ChromeDriver.Navigate().GoToUrl("http://10.10.10.79/ClearingAndSettlement/PublishReport2.aspx");

                IWebElement Date = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_CtlCalendarPickUpEOD_uiTxtCalendar"));

                IWebElement submit_publish = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_uiBtnPublish"));

                IJavaScriptExecutor js = (IJavaScriptExecutor)ChromeDriver;
                js.ExecuteScript("arguments[0].removeAttribute('readonly','readonly')", Date);

                Date.Clear();
                var tgl_eod = DateTime.Now.AddDays(-1).ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture);
                Date.SendKeys(tgl_eod);
                IWebElement generate_report = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_cbGenerateReport"));
                IWebElement replicate_report = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_cbReplicate"));
                IWebElement send_report = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_cbPublish"));
                IWebElement loading = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_UpdateProgress1"));


                //Generate Report
                generate_report.Click();
                replicate_report.Click();
                submit_publish.Submit();

                Thread.Sleep(60000);

                if (loading.Displayed)
                {
                    Thread.Sleep(60000);

                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);


                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);


                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);


                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);


                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);


                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);


                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);


                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);


                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }

                ChromeDriver.Quit();
                Thread.Sleep(60000);
                WriteToFile("Berhasil send publish report : " + DateTime.Now);
                bot.SendTextMessageAsync(chat_id, "Proses send report selesai : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Gagal Send : " + ex.Message + "\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
            }

        }
        private static string getRptImage(string reportpath)
        {
            ServerReport srv = new ServerReport();
            srv.ReportServerUrl = new Uri(System.Configuration.ConfigurationManager.ConnectionStrings["reportserver"].ConnectionString);

            srv.ReportServerCredentials = new ReportServerCredentials();
            srv.ReportPath = reportpath;
            List<ReportParameter> rp = new List<ReportParameter>();
            rp.Add(new ReportParameter("businessDateFrom", new string[] { DateTime.Now.Date.AddDays(-1).ToString("yyyy-MM-dd") }));
            //rp.Add(new ReportParameter("businessDateTo", new string[] { DateTime.Now.Date.AddDays(-1).ToString() }));

            srv.SetParameters(rp);

            byte[] result = null;
            string mimeType = "";
            string encoding = "";
            string filenameExtension = "";
            string[] streamids = null;
            Warning[] warnings = null;

            result = srv.Render("pdf", null, out mimeType, out encoding, out filenameExtension, out streamids, out warnings);

            string fileName = "Report Direksi.pdf";


            //string fileName = "Report FI Timah.pdf";

            System.IO.FileStream file_report = System.IO.File.Create(AppDomain.CurrentDomain.BaseDirectory + fileName);
            file_report.Write(result, 0, result.Length);
            file_report.Close();
            Byte[] bytes = System.IO.File.ReadAllBytes(file_report.Name);
            String file2 = "data:application/pdf;base64," + Convert.ToBase64String(bytes);

            return file_report.Name;

        }
        public static void insertContract(long chat_id, string messageFrom, string message)
        {
            var dr_contract = new DataSet1TableAdapters.ContractTableAdapter();
            var drc = new DataSet1TableAdapters.CommodityTableAdapter();
            var drs = new DataSet1TableAdapters.SubCategoryTableAdapter();
            var drt = new DataSet1TableAdapters.uspTradeFeedExceptionApproveAllTableAdapter();

            CultureInfo culture = new CultureInfo("es-ES");

            WebClient client = new WebClient();

            string[] arrmcbody = new string[50];
            if (message.Contains(" ")) ;
            {
                arrmcbody = messageFrom.Split(' ');
                if (arrmcbody[0] == "/e")
                {
                    try
                    {
                        string[] arrbody = message.Split('\n');

                        var dtc = drc.GetDataByCommodityCode(arrmcbody[4]);
                        var CommodityID = dtc[0].CommodityID;

                        Decimal SubCategoryId = dr_contract.GetData().OrderByDescending(x => x.ContractID).Where(x => x.CommodityID == CommodityID).Select(x => x.SubCategoryID).FirstOrDefault();


                        DateTime efective_start_date = Convert.ToDateTime(arrbody[0], culture);
                        DateTime start_date = Convert.ToDateTime(arrbody[0], culture);
                        DateTime end_spot = Convert.ToDateTime(arrbody[2], culture);
                        DateTime start_spot = Convert.ToDateTime(arrbody[1], culture);
                        DateTime efective_end_date = Convert.ToDateTime(arrbody[2], culture);
                        try
                        {
                            var validasi = dr_contract.GetDataByFilter(CommodityID, Convert.ToInt32(arrmcbody[12]), Convert.ToInt32(arrmcbody[8]));
                            if (validasi.Count == 0)
                            {
                                dr_contract.Insert(CommodityID,
                                          Convert.ToInt32(arrmcbody[8]),
                                          Convert.ToInt32(arrmcbody[12]),
                                          "A",
                                          efective_start_date,
                                          dtc[0].ContractSize,
                                          dtc[0].SettlementType,
                                          "",
                                          dtc[0].Unit,
                                          start_date,
                                          end_spot,
                                          start_spot,
                                          dtc[0].PEG,
                                          dtc[0].VMIRCACalType,
                                          dtc[0].SettlementFactor,
                                          dtc[0].DayRef,
                                          dtc[0].Divisor,
                                          dtc[0].MarginTender,
                                          dtc[0].MarginSpot,
                                          dtc[0].MarginRemote,
                                          dtc[0].CalSpreadRemoteMargin,
                                          dtc[0].IsKIE,
                                          "Robot System",
                                          DateTime.Now,
                                          "Robot System",
                                          DateTime.Now,
                                          efective_end_date,
                                          dtc[0].ApprovalDesc,
                                          dtc[0].HomeCurrencyID,
                                          dtc[0].CrossCurr,
                                          dtc[0].CrossCurrProduct,
                                          null,
                                          dtc[0].ActionFlag,
                                          dtc[0].TenderReqType,
                                          SubCategoryId,
                                          dtc[0].ModeK1,
                                          dtc[0].ValueK1,
                                          dtc[0].ContractRefK1,
                                          dtc[0].ModeK2,
                                          dtc[0].ValueK2,
                                          dtc[0].ContractRefK2,
                                          dtc[0].ModeD,
                                          dtc[0].ValueD,
                                          dtc[0].ContractRefD,
                                          dtc[0].ModeIM,
                                          dtc[0].PercentageSpotIM,
                                          dtc[0].PercentageRemoteIM,
                                         dtc[0].ModeFee);
                                drt.GetData(Convert.ToInt32(arrmcbody[15]), DateTime.Now.Date, "A", "ok", "Robot KBI");

                                bot.SendTextMessageAsync(chat_id, "Success insert contract \nStart date : " + arrbody[0] + "\nStart spot" + arrbody[1] + "\nEnd spot" + arrbody[2], ParseMode.Markdown);

                            }
                            else
                            {
                                bot.SendTextMessageAsync(chat_id, "The contract has been made, Please check again !", ParseMode.Markdown);
                            }

                        }
                        catch (Exception ex)
                        {
                            monitoringServices("DOP_EODTele", "E", "Service untuk insert Contract fail " + ex.Message);

                            bot.SendTextMessageAsync(chat_id, "Process insert contrct Fail : " + ex.Message, ParseMode.Markdown);
                        }
                    }
                    catch (Exception ex)
                    {
                        monitoringServices("DOP_EODTele", "E", "Service untuk insert Contract fail " + ex.Message);

                        bot.SendTextMessageAsync(chat_id, "Process insert contrct Fail : " + ex.Message, ParseMode.Markdown);
                    }
                }
            }
        }
        public static void getdataPaln(long chat_id)
        {
            try
            {
                var dr = new DataSet1TableAdapters.iDfsFasPalnTableAdapter();
                var dt = dr.GetData();
                List<string> data = new List<string>();
                if (dt.Count != 0)
                {
                    foreach (var item in dt)
                    {
                        data.Add("BussinessDate : " + item.BusinessDate.ToString("dd/MM/yyyy") + "\nSegDepositUSD : " + item.PALN.ToString("#,##0.##") + "\nClearing Member Name : " + item.Name);
                    }
                }
                var msg = String.Join("\n\n", data);
                bot.SendTextMessageAsync(chat_id, msg, ParseMode.Markdown);

            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Get data paln fail :" + ex.Message, ParseMode.Markdown);
            }
        }
        public static async void rptShortageDailyDKA(long chat_id)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Generate start " + DateTime.Now.ToString("HH:mm:ss"), ParseMode.Markdown);

                string path = getReportSSRSPDFExcel("RptDailyShortageSummaryForWA", "&businessDate=" + DateTime.Now.ToString("yyyy-MM-dd"));
                sendFileTelegram("-1001671146559", path);

            }
            catch (Exception x)
            {
                bot.SendTextMessageAsync(chat_id, "Generate failed " + x.Message + " " + DateTime.Now.ToString("HH:mm:ss"), ParseMode.Markdown);
            }
        }
        public static void insertIrca(long chat_id, string path)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Process insert IRCA start " + DateTime.Now.ToString("HH:mm:ss"), ParseMode.Markdown);
                List<irca_parameter> data_excel = new List<irca_parameter>();
                string con = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=Yes;';";
                using (OleDbConnection connection = new OleDbConnection(con))
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand("Select * From [Sheet1$]", connection);
                    using (OleDbDataReader dr = command.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            var data_awal = new irca_parameter
                            {
                                pair = dr[0].ToString(),
                                year = (dr[1].ToString()),
                                day = (dr[2].ToString()),
                            };
                            data_excel.Add(data_awal);
                        }
                    }
                }
                var ta = new DataSet1TableAdapters.irca_parameterTableAdapter();
                var dt = ta.GetData();
                foreach (var old in data_excel)
                {
                    foreach (var neww in dt)
                    {
                        if (old.pair == neww.pair)
                        {
                            Decimal value_old = 0;
                            if (old.day != "-")
                            {
                                value_old = (Decimal.Parse(old.day.ToString(), System.Globalization.NumberStyles.Any));
                            }
                            var olah_data = new irca_parameter_new
                            {
                                pair = neww.pair,
                                old_value = value_old,
                                parameter = neww.parameter,
                                new_value = value_old * neww.parameter
                            };

                            new_data.Add(olah_data);
                        }
                    }
                }
                var ta_irp = new DataSet1TableAdapters.irca_productTableAdapter();
                var tac = new DataSet1TableAdapters.CommodityTableAdapter();
                var tar = new DataSet1TableAdapters.IRCATableAdapter();
                var dt_irp = ta_irp.GetData();
                List<irca_parameter_final> data_insert = new List<irca_parameter_final>();
                string msg = "";
                foreach (var old in dt_irp)
                {
                    foreach (var neww in new_data)
                    {
                        if (old.pair == neww.pair)
                        {
                            var insert = new irca_parameter_final
                            {
                                product = old.product,
                                date = DateTime.Now,
                                irca = neww.new_value
                            };
                            var dr = tac.GetDataByComCode(old.product);
                            data_insert.Add(insert);
                            var startdate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                            var dri = tar.GetDataByFilter(dr[0].CommodityID, startdate);
                            if (dri.Count != 0)
                            {
                                var dar = tar.GetDataById(dri[0].IRCAID);
                                dar[0].CommodityID = dr[0].CommodityID;
                                dar[0].EffectiveStartDate = startdate;
                                dar[0].ApprovalStatus = "A";
                                dar[0].IRCAValue = neww.new_value;
                                dar[0].CreatedBy = "System";
                                dar[0].LastUpdatedBy = "System";
                                dar[0].LastUpdatedDate = DateTime.Now;
                                dar[0].EffectiveEndDate = Convert.ToDateTime(DateTime.Now.Year + "/" + DateTime.Now.Month + "/" + DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month));
                                dar[0].ApprovalDesc = "IRCA " + DateTime.Now.ToString("MMMM") + " " + DateTime.Now.Year;
                                tar.Update(dar);
                                msg = "the data already exists, successfully updated " + DateTime.Now.ToString("HH:mm:ss");
                            }
                            else
                            {
                                tar.InsertQuery(dr[0].CommodityID, startdate, "A", neww.new_value, "System", DateTime.Now, "System", DateTime.Now, Convert.ToDateTime(DateTime.Now.Year + "/" + DateTime.Now.Month + "/" + DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month)), "IRCA " + DateTime.Now.ToString("MMMM") + " " + DateTime.Now.Year, null, null);
                                msg = "Process insert IRCA success " + DateTime.Now.ToString("HH:mm:ss");
                            }
                        }
                    }
                }
                StringBuilder sb = new StringBuilder();

                sb.AppendLine("product;date;irca");

                foreach (var lvpn in data_insert)
                {
                    string line = lvpn.product + ";" + lvpn.date + ";" + lvpn.irca;
                    sb.AppendLine(line);
                }

                string pathCsv = AppDomain.CurrentDomain.BaseDirectory + "\\irca.csv";
                System.IO.File.WriteAllText(pathCsv, sb.ToString());
                sendFileTelegram(chat_id.ToString(), pathCsv);
                bot.SendTextMessageAsync(chat_id, msg, ParseMode.Markdown);

            }
            catch (Exception x)
            {
                bot.SendTextMessageAsync(chat_id, "Process insert IRCA failed " + x.Message + " " + DateTime.Now.ToString("HH:mm:ss"), ParseMode.Markdown);
            }
        }
        public static void pengecekanReportTerkirim(long chat_id)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Proses pengecekan report start : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                IWebDriver ChromeDriver = new ChromeDriver();
                ChromeDriver.Manage().Window.Maximize();
                ChromeDriver.Navigate().GoToUrl("https://10.10.10.2:4444/");
                IWebElement advanced = ChromeDriver.FindElement(By.Id("details-button"));
                IWebElement proced = ChromeDriver.FindElement(By.Id("proceed-link"));
                advanced.Click();
                proced.Click();
                Thread.Sleep(2000);
                IWebElement username = ChromeDriver.FindElement(By.Id("ELEMENT_login_username"));
                IWebElement password = ChromeDriver.FindElement(By.Id("ELEMENT_login_password"));
                IWebElement submit_login = ChromeDriver.FindElement(By.Id("ELEMENT_login_button"));
                username.SendKeys("admin");
                password.SendKeys("Sophos123!@");
                submit_login.Click();
                Thread.Sleep(10000);
                ((IJavaScriptExecutor)ChromeDriver).ExecuteScript("window.open('https://10.10.10.2:4444/qm/?lang=english');");
                ChromeDriver.SwitchTo().Window(ChromeDriver.WindowHandles[1]);

                new Actions(ChromeDriver).SendKeys(ChromeDriver.FindElement(By.TagName("html")), Keys.Control).SendKeys(ChromeDriver.FindElement(By.TagName("html")), Keys.NumberPad2);
                Thread.Sleep(5000);
                ChromeDriver.FindElement(By.XPath("//div[@class='menuitem'][contains(.,'SMTP Log')]")).Click();
                Thread.Sleep(5000);
                //((ITakesScreenshot)ChromeDriver).GetScreenshot().SaveAsFile("report.png");
                ((IJavaScriptExecutor)ChromeDriver).ExecuteScript("document.body.style.zoom = '0.7'", string.Empty);

                Thread.Sleep(500);
                ((IJavaScriptExecutor)ChromeDriver).ExecuteScript("window.scrollBy(0, 150)", string.Empty);
                var x = ((ITakesScreenshot)ChromeDriver).GetScreenshot();
                x.SaveAsFile(AppDomain.CurrentDomain.BaseDirectory + "\\report1.png");
                Thread.Sleep(500);
                ((IJavaScriptExecutor)ChromeDriver).ExecuteScript("window.scrollBy(0, 450)", string.Empty);
                Thread.Sleep(500);
                x = ((ITakesScreenshot)ChromeDriver).GetScreenshot();
                //x.SaveAsFile(@"D:\report2.png");
                x.SaveAsFile(AppDomain.CurrentDomain.BaseDirectory + "\\report2.png");
                //.SaveAsFile("report.png");
                ChromeDriver.Quit();
                sendFileTelegram("-1001671146559", AppDomain.CurrentDomain.BaseDirectory + "\\report1.png");
                sendFileTelegram("-1001671146559", AppDomain.CurrentDomain.BaseDirectory + "\\report2.png");

            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Proses pengecekan report failed : " + ex.Message + " " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
            }
        }
        public static void generetaRptCluster(long chat_id, string body)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Proses get report start : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

                string month = body;
                var date = validateMonth(month).Split(' ');

                string path = getReportSSRSPDFExcel("RptClearingMemberMonthly", "&businessDateFrom=" + date[0] + "&businessDateTo=" + date[1] + "&startDate=" + date[0] + "&endDate=" + date[1]);
                sendFileTelegram("-1001671146559", path);
                bot.SendTextMessageAsync(chat_id, "Proses get report finish : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Proses get report fail : " + ex.Message + "\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
            }
        }
        public static void insert_paln_sfi_1(long chat_id, string body)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Proses read file pdf start : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                IronOcr.Installation.LicenseKey = "IRONOCR.PTKLIRINGBERJANGKAINDONESIA.IRO211213.9250.23127.312112-1EA71407D8-HWNZXEAJ3YDY3-N2BU5BL3WRB6-YYXFD7XQVLTB-YLR4RKFSW22L-F7OLHD7TWYX3-MUPLUQ-LVPA4EVNX2WIEA-PROFESSIONAL.SUB-DNMTTQ.RENEW.SUPPORT.13.DEC.2022";
                var Ocr = new IronTesseract();
                using (var input = new OcrInput())
                {
                    input.AddPdf(body);
                    // We can also select specific PDF page numnbers to OCR
                    var Result = Ocr.Read(input);
                    var splitresult = Result.Text.Split(new string[] { "\r\nTotal Equity " }, StringSplitOptions.None);
                    var hasil = splitresult[1].Split(' ')[0];
                    System.IO.File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "si1.txt", hasil);
                    bot.SendTextMessageAsync(chat_id, "Proses get total equity SFI : " + hasil + " " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                    File.Delete(body);
                    // 1 page for every page of the PDF
                }
            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Proses get report fail : " + ex.Message + "\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
            }
        }
        public static void insert_paln_sfi_2(long chat_id, string body)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Proses read file pdf start : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                IronOcr.Installation.LicenseKey = "IRONOCR.PTKLIRINGBERJANGKAINDONESIA.IRO211213.9250.23127.312112-1EA71407D8-HWNZXEAJ3YDY3-N2BU5BL3WRB6-YYXFD7XQVLTB-YLR4RKFSW22L-F7OLHD7TWYX3-MUPLUQ-LVPA4EVNX2WIEA-PROFESSIONAL.SUB-DNMTTQ.RENEW.SUPPORT.13.DEC.2022";
                var Ocr = new IronTesseract();
                using (var input = new OcrInput())
                {
                    input.AddPdf(body);
                    // We can also select specific PDF page numnbers to OCR
                    var Result = Ocr.Read(input);
                    var splitresult = Result.Text.Split(new string[] { "\r\nTotal Equity " }, StringSplitOptions.None);
                    var hasil = splitresult[1].Split(' ')[0];
                    System.IO.File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "si2.txt", hasil);
                    bot.SendTextMessageAsync(chat_id, "Proses get total equity SFI : " + hasil + " " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                    File.Delete(body);
                    // 1 page for every page of the PDF
                }
            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Proses get report fail : " + ex.Message + "\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
            }
        }
        public static void insert_paln_pg_1(long chat_id, string body)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Proses read file pdf start : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                IronOcr.Installation.LicenseKey = "IRONOCR.PTKLIRINGBERJANGKAINDONESIA.IRO211213.9250.23127.312112-1EA71407D8-HWNZXEAJ3YDY3-N2BU5BL3WRB6-YYXFD7XQVLTB-YLR4RKFSW22L-F7OLHD7TWYX3-MUPLUQ-LVPA4EVNX2WIEA-PROFESSIONAL.SUB-DNMTTQ.RENEW.SUPPORT.13.DEC.2022";
                var Ocr = new IronTesseract();
                using (var input = new OcrInput())
                {
                    input.AddPdf(body);
                    // We can also select specific PDF page numnbers to OCR
                    var Result = Ocr.Read(input);
                    var splitresult = Result.Text.Split(new string[] { "\r\nTotal Equity " }, StringSplitOptions.None);
                    var hasil = splitresult[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0];
                    System.IO.File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "pg1.txt", hasil);
                    bot.SendTextMessageAsync(chat_id, "Proses get total equity PG : " + hasil + " " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                    File.Delete(body);
                    // 1 page for every page of the PDF
                }
            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Proses get report fail : " + ex.Message + "\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
            }
        }
        public static void insert_paln_pg_2(long chat_id, string body)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Proses read file pdf start : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                IronOcr.Installation.LicenseKey = "IRONOCR.PTKLIRINGBERJANGKAINDONESIA.IRO211213.9250.23127.312112-1EA71407D8-HWNZXEAJ3YDY3-N2BU5BL3WRB6-YYXFD7XQVLTB-YLR4RKFSW22L-F7OLHD7TWYX3-MUPLUQ-LVPA4EVNX2WIEA-PROFESSIONAL.SUB-DNMTTQ.RENEW.SUPPORT.13.DEC.2022";
                var Ocr = new IronTesseract();
                using (var input = new OcrInput())
                {
                    input.AddPdf(body);
                    // We can also select specific PDF page numnbers to OCR
                    var Result = Ocr.Read(input);
                    var x = Result.Text.Split(' ');
                    var hasil = x[(x.Length - 1)].Replace("$", "");
                    System.IO.File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "pg2.txt", hasil);
                    bot.SendTextMessageAsync(chat_id, "Proses get total equity PG ALPACA : " + hasil + " " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                    File.Delete(body);
                    // 1 page for every page of the PDF
                }
            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Proses get report fail : " + ex.Message + "\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
            }
        }
        public static void insert_paln_pg_3(long chat_id, string body)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Proses read file pdf start : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                IronOcr.Installation.LicenseKey = "IRONOCR.PTKLIRINGBERJANGKAINDONESIA.IRO211213.9250.23127.312112-1EA71407D8-HWNZXEAJ3YDY3-N2BU5BL3WRB6-YYXFD7XQVLTB-YLR4RKFSW22L-F7OLHD7TWYX3-MUPLUQ-LVPA4EVNX2WIEA-PROFESSIONAL.SUB-DNMTTQ.RENEW.SUPPORT.13.DEC.2022";
                var Ocr = new IronTesseract();
                using (var input = new OcrInput())
                {
                    input.AddPdf(body, "71001abc");
                    // We can also select specific PDF page numnbers to OCR
                    var Result = Ocr.Read(input);
                    var splitresult = Result.Text.Split(new string[] { "\r\nTotal Equity " }, StringSplitOptions.None);
                    var hasil = splitresult[1].Split(' ')[0];
                    System.IO.File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "pg3.txt", hasil);
                    bot.SendTextMessageAsync(chat_id, "Proses get total equity PG : " + hasil + " " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                    File.Delete(body);
                    // 1 page for every page of the PDF
                }
            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Proses get report fail : " + ex.Message + "\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
            }
        }
        public static void insert_data_paln(long chat_id)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Proses insert paln start : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                var dr = new DataSet1TableAdapters.iDfsFasPaln1TableAdapter();
                var dt_strait = dr.GetDataByDate(DateTime.Now, 115);
                var dt_pg = dr.GetDataByDate(DateTime.Now, 119);
                var dt_philip = dr.GetDataByDate(DateTime.Now, 121);
                if (dt_strait.Count == 0)
                {
                    string si1 = AppDomain.CurrentDomain.BaseDirectory + "\\si1.txt";
                    string si2 = AppDomain.CurrentDomain.BaseDirectory + "\\si2.txt";
                    string[] text1 = File.ReadAllLines(si1);
                    string[] text2 = File.ReadAllLines(si2);
                    var hasil = Convert.ToDecimal(text1[0].Replace(",", "").Replace(".", ",")) + Convert.ToDecimal(text2[0].Replace(",", "").Replace(".", ","));
                    bot.SendTextMessageAsync(chat_id, "Strait 1 = " + text1[0] + "\nStrait 2 = " + text2[0] + "\nTotal = " + hasil);

                    dr.Insert(115, DateTime.Now.Date, hasil);
                    //msg += "PT STRAITS FUTURES INDONESIA : " + value;
                }
                if (dt_pg.Count == 0)
                {
                    string pg1 = AppDomain.CurrentDomain.BaseDirectory + "\\pg1.txt";
                    string pg2 = AppDomain.CurrentDomain.BaseDirectory + "\\pg2.txt";
                    string pg3 = AppDomain.CurrentDomain.BaseDirectory + "\\pg3.txt";
                    string[] text1 = File.ReadAllLines(pg1);
                    string[] text2 = File.ReadAllLines(pg2);
                    string[] text3 = File.ReadAllLines(pg3);
                    var hasil = Convert.ToDecimal(text1[0].Replace(",", "").Replace(".", ",")) + Convert.ToDecimal(text2[0].Replace(",", "").Replace(".", ",")) + Convert.ToDecimal(text3[0].Replace(",", "").Replace(".", ","));
                    bot.SendTextMessageAsync(chat_id, "PG 1 = " + text1[0] + "\nPG 2 = " + text2[0] + "\nALPACA PG = " + text3[0] + "\nTotal = " + hasil);

                    dr.Insert(119, DateTime.Now.Date, hasil);
                    //msg += "\nPT PG BERJANGKA : " + value;
                }
                bot.SendTextMessageAsync(chat_id, "Proses insert paln sucess : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
            }
            catch (Exception ex)
            {
                bot.SendTextMessageAsync(chat_id, "Proses insert paln fail : " + ex.Message + "\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
            }
        }
        public static void gettotalpaln(long chat_id, string msg)
        {
            try
            {
                bot.SendTextMessageAsync(chat_id, "Proses get total PALN start : " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                var dr = new DataSet1TableAdapters.fee_palnTableAdapter();
                var date = msg.Split(new string[] { "\n" }, StringSplitOptions.None);
                var dt = dr.GetDataByDate(Convert.ToDateTime(date[0]), Convert.ToDateTime(date[1]));

                var dr_kurs = new DataSet1TableAdapters.RateJisdorTableAdapter();
                var rate = dr_kurs.GetDataByDate(Convert.ToDateTime(date[1]));

                var total = dt[0].totalFee * rate[0].Rate;
                bot.SendTextMessageAsync(chat_id, "Total Fee USD : " + dt[0].totalFee.ToString("#,##0.####") + "\nKurs Tanggal " + date[1] + " : " + rate[0].Rate.ToString("#,##0.##") + "\nTotal : " + total.ToString("#,##0.##") + "\n\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);

            }
            catch (Exception xx)
            {
                bot.SendTextMessageAsync(chat_id, "Service untuk get data total PALN fail " + xx.Message + " " + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                monitoringServices("DOP_EODTele", "E", "Service untuk get data total PALN fail " + xx.Message);
            }
        }
        public class data_shortage_direksi
        {
            public string nama { get; set; }
            public decimal nominal { get; set; }
        }
        public static bool IsHoliday(DateTime inputDate)
        {
            bool isHoliday;
            var ta = new DataSet1TableAdapters.QueriesTableAdapter();

            try
            {
                isHoliday = ta.IsHoliday(inputDate).Value;

                return isHoliday;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message, ex);
            }
        }
        protected override void OnStop()
        {
            WriteToFile("Service is started at " + DateTime.Now);
            bot.StopReceiving();
        }
        public static void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
        private static string SendTelegram(string chatId, string message)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls12;

            var client = new RestClient("https://api.telegram.org/bot2144239635:AAFjcfn_GdHP4OkzzZomaZt4XbwpHDGyR-U/sendMessage?chat_id=" + chatId + "&text=" + message);
            client.Timeout = -1;
            var requestWa = new RestRequest(Method.GET);
            IRestResponse responseWa = client.Execute(requestWa);
            return responseWa.Content;
        }
        private static void sendFileTelegram(string chatId, string body)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls12;
            var client = new RestClient("https://api.telegram.org/bot2144239635:AAFjcfn_GdHP4OkzzZomaZt4XbwpHDGyR-U/sendDocument");
            client.Timeout = -1;
            var requestWa = new RestRequest(Method.POST);
            requestWa.AddHeader("Content-Type", "multipart/form-data");
            requestWa.AddParameter("chat_id", chatId);
            requestWa.AddFile("document", body);
            IRestResponse responseWa = client.Execute(requestWa);
            Console.WriteLine(responseWa.Content);
        }
        public static string validateMonth(string month)
        {
            if (month == "januari")
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, 1, 1);
                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
                return firstDayOfMonth.ToString("yyyy-MM-dd") + " " + lastDayOfMonth.ToString("yyyy-MM-dd");
            }
            if (month == "februari")
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, 2, 1);
                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
                return firstDayOfMonth.ToString("yyyy-MM-dd") + " " + lastDayOfMonth.ToString("yyyy-MM-dd");
            }
            if (month == "maret")
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, 3, 1);
                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
                return firstDayOfMonth.ToString("yyyy-MM-dd") + " " + lastDayOfMonth.ToString("yyyy-MM-dd");
            }
            if (month == "april")
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, 4, 1);
                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
                return firstDayOfMonth.ToString("yyyy-MM-dd") + " " + lastDayOfMonth.ToString("yyyy-MM-dd");
            }
            if (month == "mei")
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, 5, 1);
                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
                return firstDayOfMonth.ToString("yyyy-MM-dd") + " " + lastDayOfMonth.ToString("yyyy-MM-dd");
            }
            if (month == "juni")
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, 6, 1);
                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
                return firstDayOfMonth.ToString("yyyy-MM-dd") + " " + lastDayOfMonth.ToString("yyyy-MM-dd");
            }
            if (month == "juli")
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, 7, 1);
                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
                return firstDayOfMonth.ToString("yyyy-MM-dd") + " " + lastDayOfMonth.ToString("yyyy-MM-dd");
            }
            if (month == "agustus")
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, 8, 1);
                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
                return firstDayOfMonth.ToString("yyyy-MM-dd") + " " + lastDayOfMonth.ToString("yyyy-MM-dd");
            }
            if (month == "september")
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, 9, 1);
                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
                return firstDayOfMonth.ToString("yyyy-MM-dd") + " " + lastDayOfMonth.ToString("yyyy-MM-dd");
            }
            if (month == "oktober")
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, 10, 1);
                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
                return firstDayOfMonth.ToString("yyyy-MM-dd") + " " + lastDayOfMonth.ToString("yyyy-MM-dd");
            }
            if (month == "november")
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, 11, 1);
                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
                return firstDayOfMonth.ToString("yyyy-MM-dd") + " " + lastDayOfMonth.ToString("yyyy-MM-dd");
            }
            else
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, 12, 1);
                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
                return firstDayOfMonth.ToString("yyyy-MM-dd") + " " + lastDayOfMonth.ToString("yyyy-MM-dd");
            }

        }
        public static string getReportSSRSPDFExcel(string filename, string parameter)
        {
            try
            {
                string ori_url = System.Configuration.ConfigurationManager.ConnectionStrings["reportserver"].ConnectionString;
                string url = ori_url + "?/SKDReport/" + filename + "&rs:Command=Render&rs:Format=EXCELOPENXML&rc:OutputFormat=XLS" + parameter;

                System.Net.HttpWebRequest Req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
                Req.Credentials = new NetworkCredential("administrator", "Indonesia@2022");
                Req.Method = "GET";
                Req.Timeout = 1200000;
                string path = AppDomain.CurrentDomain.BaseDirectory + "\\" + filename + ".xlsx";

                System.Net.WebResponse objResponse = Req.GetResponse();
                System.IO.FileStream fs = new System.IO.FileStream(path, System.IO.FileMode.Create);
                System.IO.Stream stream = objResponse.GetResponseStream();

                byte[] buf = new byte[1024];
                int len = stream.Read(buf, 0, 1024);
                while (len > 0)
                {
                    fs.Write(buf, 0, len);
                    len = stream.Read(buf, 0, 1024);
                }
                stream.Close();
                fs.Close();
                return path;
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        public class irca_parameter
        {
            public string pair { get; set; }
            public string year { get; set; }
            public string day { get; set; }
        }
        public class irca_parameter_new
        {
            public String pair { get; set; }
            public int parameter { get; set; }
            public Decimal old_value { get; set; }
            public Decimal new_value { get; set; }
        }
        public class irca_parameter_final
        {
            public string product { get; set; }
            public DateTime date { get; set; }
            public Decimal irca { get; set; }
        }
        private static string monitoringServices(string servicename, string status, string desc)
        {
            string jsonString = "{" +
                     "\"name\" : \"" + servicename + "\"," +
                     "\"logstatus\": \"" + status + "\"," +
                     "\"logdesc\":\"" + desc + "\"," +
                     "}";
            var client = new RestClient("https://apiservicekbi.azurewebsites.net/api/ServiceStatus");

            RestRequest requestWa = new RestRequest("https://apiservicekbi.azurewebsites.net/api/ServiceStatus", Method.POST);
            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("data", jsonString);
            var responseWa = client.ExecutePostAsync(requestWa);
            return (responseWa.Result.Content);
        }
    }
}
