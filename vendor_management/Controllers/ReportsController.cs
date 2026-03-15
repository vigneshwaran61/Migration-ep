using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using BusinessAnalysisLayer;
using BusinessObject;
using iText.Html2pdf;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;

namespace SYS_VENDOR_MGMT.Controllers
{
    public class ReportsController : Controller
    {
        private readonly IConfiguration _config;
        private readonly IWebHostEnvironment _env;
        private readonly SoapSequenceClient _seqClient;

        public ReportsController(IConfiguration config, IWebHostEnvironment env, IHttpClientFactory httpClientFactory)
        {
            _config = config;
            _env = env;
            _seqClient = new SoapSequenceClient(httpClientFactory, config);
        }

        // ── GET: /Reports/MonthlyPaymentReport ───────────────────────────────
        public IActionResult MonthlyPaymentReport()
        {
            var vm = new MonthlyPaymentReportViewModel();
            try
            {
                BindMonths(vm);
            }
            catch (Exception ex)
            {
                ExceptionUtility.LogException(ex, ex.Source ?? nameof(MonthlyPaymentReport));
            }
            return View(vm);
        }

        // ── AJAX GET: cascade Vendor dropdown ────────────────────────────────
        [HttpGet]
        public IActionResult BindVendors(string month)
        {
            try
            {
                string Err = string.Empty;
                var objRateRevisionBal = new RateRevisionBal();
                string[] mnth = month.Split(',');
                DataTable dt = objRateRevisionBal.VIEW_RATE_REVISION_MST("", ref Err);
                DataView dv = dt.DefaultView;
                dv.RowFilter = $"f_Month='{mnth[0].Trim()}' and f_Year={mnth[1].Trim()}";
                string[] columns = { "f_VendorCode", "f_VendorName" };
                DataTable uniqueCols = dv.ToTable(true, columns);

                var result = new System.Collections.Generic.List<object>();
                foreach (DataRow row in uniqueCols.Rows)
                    result.Add(new { value = row["f_VendorCode"].ToString(), text = row["f_VendorName"].ToString() });
                return Json(result);
            }
            catch (Exception ex)
            {
                ExceptionUtility.LogException(ex, ex.Source ?? nameof(BindVendors));
                return Json(new System.Collections.Generic.List<object>());
            }
        }

        // ── AJAX GET: cascade PO Number dropdown ─────────────────────────────
        [HttpGet]
        public IActionResult BindPONumbers(string month, string vendorCode)
        {
            try
            {
                string Err = string.Empty;
                var objRateRevisionBal = new RateRevisionBal();
                string[] mnth = month.Split(',');
                DataTable dt = objRateRevisionBal.VIEW_RATE_REVISION_MST("", ref Err);
                DataView dv = dt.DefaultView;
                dv.RowFilter = $"f_Month='{mnth[0].Trim()}' and f_Year={mnth[1].Trim()} and f_VendorCode='{vendorCode}'";

                var result = new System.Collections.Generic.List<object>();
                foreach (DataRowView row in dv)
                    result.Add(new { value = row["f_RateCode"].ToString(), text = row["f_PONumber"].ToString() });
                return Json(result);
            }
            catch (Exception ex)
            {
                ExceptionUtility.LogException(ex, ex.Source ?? nameof(BindPONumbers));
                return Json(new System.Collections.Generic.List<object>());
            }
        }

        // ── POST: Search ──────────────────────────────────────────────────────
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Search(ReportSearchModel model)
        {
            var vm = new MonthlyPaymentReportViewModel();
            try
            {
                BindMonths(vm);
                vm.Search = model;

                if (string.IsNullOrEmpty(model.SelectedVendorCode))
                {
                    vm.ErrorMessage = "Please select vendor.";
                    return View("MonthlyPaymentReport", vm);
                }
                if (string.IsNullOrEmpty(model.SelectedPOValue))
                {
                    vm.ErrorMessage = "Please select PO.";
                    return View("MonthlyPaymentReport", vm);
                }
                if (string.IsNullOrEmpty(model.SelectedMonth))
                {
                    vm.ErrorMessage = "Please select Month.";
                    return View("MonthlyPaymentReport", vm);
                }

                string Err = string.Empty;
                var objPOProcessBAL = new POProcessBAL();
                DataSet dsReport = objPOProcessBAL.GET_RPT_FINAL_MONTHLY_PAYMENT_SLIP(
                    model.SelectedPOText, model.SelectedPOValue, model.SelectedMonth, ref Err);

                if (dsReport != null && dsReport.Tables.Count >= 3)
                {
                    vm.ShowDetails = true;
                    vm.Details = BuildDetailsModel(dsReport.Tables[0]);
                    vm.GridRows = dsReport.Tables[1];
                    if (dsReport.Tables.Count > 3)
                        BuildTotalModel(dsReport.Tables[3], vm, model);
                    vm.Header = $"Final Monthly Payment Slip For {vm.Details.VendorName} ( {model.SelectedMonth} )";

                    // Check covering letter availability
                    string Err2 = string.Empty;
                    var objPOTaxDetailsBAL = new POTaxDetailsBAL();
                    vm.ShowCoveringBtn = objPOTaxDetailsBAL.CHECK_SVM_COVERING(
                        model.SelectedPOValue, vm.Details.VendorCode.Trim(), model.SelectedMonth, ref Err2);
                }
                else
                {
                    vm.ErrorMessage = "PO Process Not Done";
                }
            }
            catch (Exception ex)
            {
                ExceptionUtility.LogException(ex, ex.Source ?? nameof(Search));
                vm.ErrorMessage = ex.Message;
            }
            return View("MonthlyPaymentReport", vm);
        }

        // ── POST: Generate ────────────────────────────────────────────────────
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Generate(ReportSearchModel model)
        {
            try
            {
                UserLogin? objUserLogin = HttpContext.Session.Get<UserLogin>("UserLogin");
                if (objUserLogin == null)
                    return Json(new { success = false, message = "Session expired." });

                string Err = string.Empty;
                var objPOProcessBAL = new POProcessBAL();
                var objPOProcessBO  = new POProcessBO();

                objPOProcessBO.PONo            = model.SelectedPOText;
                objPOProcessBO.RateCode        = model.SelectedPOValue;
                objPOProcessBO.MonthName       = model.SelectedMonth;
                objPOProcessBO.CreatedType     = objUserLogin.EntityType;
                objPOProcessBO.CreatedID       = objUserLogin.EntityID;
                objPOProcessBO.CreatedBy       = objUserLogin.CreatedBy;
                objPOProcessBO.CreatedByUName  = objUserLogin.CreatedByUname;
                objPOProcessBO.CreatedIP       = HttpContext.Connection.RemoteIpAddress?.ToString() ?? "";
                objPOProcessBO.SessionID       = Convert.ToInt32(objUserLogin.SessionId);
                objPOProcessBO.FormCode        = objUserLogin.FormCode;
                objPOProcessBO.FormName        = objUserLogin.FormName;
                objPOProcessBO.VendorCode      = model.SelectedVendorCode;

                // Get sequence numbers via SOAP (replaces Vendor_SEQ.VENDOR_SEQ ASMX proxy)
                objPOProcessBO.Sequence = await _seqClient.GetSvmPaymentGeneratedDtlSeqAsync();
                int result = objPOProcessBAL.SAVE_PAYMENT_GENERATED(objPOProcessBO, ref Err);

                if (result != 0)
                {
                    objPOProcessBO.Sequence = await _seqClient.GetSvmVendorNetpayDtlSeqAsync();
                    objPOProcessBAL.SAVE_VENDOR_NET_PAYMENT(objPOProcessBO, model.DebitReason?.Trim() ?? "", ref Err);
                }

                if (!string.IsNullOrEmpty(Err))
                    return Json(new { success = false, message = Err });

                // Check covering
                string Err2 = string.Empty;
                var objPOTaxDetailsBAL = new POTaxDetailsBAL();
                bool hasCovering = objPOTaxDetailsBAL.CHECK_SVM_COVERING(
                    model.SelectedPOValue, model.SelectedVendorCode.Trim(), model.SelectedMonth, ref Err2);

                return Json(new { success = true, message = "Generated successfully", hasCovering });
            }
            catch (Exception ex)
            {
                ExceptionUtility.LogException(ex, ex.Source ?? nameof(Generate));
                return Json(new { success = false, message = ex.Message });
            }
        }

        // ── POST: Export Invoice PDF ──────────────────────────────────────────
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult ExportInvoice(ReportSearchModel model)
        {
            try
            {
                string Err = string.Empty;
                var objPOProcessBAL = new POProcessBAL();
                DataSet dsReport = objPOProcessBAL.GET_RPT_FINAL_MONTHLY_PAYMENT_SLIP(
                    model.SelectedPOText, model.SelectedPOValue, model.SelectedMonth, ref Err);

                DataTable dtDetails = dsReport.Tables[0];
                DataTable dtGrid    = dsReport.Tables[1];
                DataTable dtTotal   = dsReport.Tables[3];

                var sb = new StringBuilder();
                if (dtDetails.Rows.Count > 0)
                {
                    sb.Append("<html><head><title></title></head><body>");
                    sb.Append(GenerateStringForPO(dtDetails, dtGrid, dtTotal, model));
                    sb.Append("</body></html>");
                }

                string legcode     = dtDetails.Rows[0]["LegacyCode"].ToString()!;
                string pdffilename = $"{legcode}_{model.SelectedMonth}_{DateTime.Now.Ticks}.pdf";
                string reportDir   = Path.Combine(_env.WebRootPath, "Reports");
                Directory.CreateDirectory(reportDir);
                string pdfPath     = Path.Combine(reportDir, pdffilename);

                byte[] pdfBytes = ConvertHtmlToPdfLedger(sb.ToString(), pdfPath);

                // Log invoice view
                InvoiceViewCount("INV", pdffilename, model);

                return File(pdfBytes, "application/pdf", pdffilename);
            }
            catch (Exception ex)
            {
                ExceptionUtility.LogException(ex, ex.Source ?? nameof(ExportInvoice));
                TempData["ErrorMsg"] = ex.Message;
                return RedirectToAction(nameof(MonthlyPaymentReport));
            }
        }

        // ── POST: Export Covering Letter PDF ─────────────────────────────────
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult ExportCovering(ReportSearchModel model)
        {
            try
            {
                string Err = string.Empty;
                var objPOProcessBAL = new POProcessBAL();
                DataSet dsReport = objPOProcessBAL.GET_RPT_FINAL_MONTHLY_PAYMENT_SLIP(
                    model.SelectedPOText, model.SelectedPOValue, model.SelectedMonth, ref Err);

                DataTable dtPOTax  = objPOProcessBAL.GET_POTAXDETAILS(model.SelectedPOValue, model.SelectedMonth, ref Err);
                DataTable dtDetails = dsReport.Tables[0];
                DataTable dtGrid    = dsReport.Tables[1];
                DataTable dtTotal   = dsReport.Tables[3];

                var sb = new StringBuilder();
                if (dtDetails.Rows.Count > 0)
                {
                    sb.Append("<html><head><title></title></head><body>");
                    sb.Append(GenerateCoverLetter(dtDetails, dtGrid, dtTotal, dtPOTax, model));
                    sb.Append("</body></html>");
                }

                string legcode     = dtDetails.Rows[0]["Legacycode"].ToString()!;
                string pdffilename = $"{legcode}_{model.SelectedMonth}_Cover_{DateTime.Now.Ticks}.pdf";
                string reportDir   = Path.Combine(_env.WebRootPath, "Reports");
                Directory.CreateDirectory(reportDir);
                string pdfPath     = Path.Combine(reportDir, pdffilename);

                byte[] pdfBytes = ConvertHtmlToPdfA4(sb.ToString(), pdfPath);

                InvoiceViewCount("COV", pdffilename, model);

                return File(pdfBytes, "application/pdf", pdffilename);
            }
            catch (Exception ex)
            {
                ExceptionUtility.LogException(ex, ex.Source ?? nameof(ExportCovering));
                TempData["ErrorMsg"] = ex.Message;
                return RedirectToAction(nameof(MonthlyPaymentReport));
            }
        }

        // ═══════════════════════════════════════════════════════════════════════
        // Private Helpers — ported verbatim from MonthlyPaymentReport_Admin.aspx.cs
        // ═══════════════════════════════════════════════════════════════════════

        private void BindMonths(MonthlyPaymentReportViewModel vm)
        {
            string Err = string.Empty;
            var objRateRevisionBal = new RateRevisionBal();
            DataTable dt = objRateRevisionBal.GET_DISTINCT_MONTH("", ref Err);
            vm.Months = new System.Collections.Generic.List<SelectListItem>();
            vm.Months.Add(new SelectListItem(" -Select- ", ""));
            if (dt.Rows.Count > 0)
                foreach (DataRow row in dt.Rows)
                    vm.Months.Add(new SelectListItem(row["MonthName"].ToString(), row["MonthName"].ToString()));
        }

        private ReportDetailsModel BuildDetailsModel(DataTable dt)
        {
            if (dt.Rows.Count == 0) return new ReportDetailsModel();
            return new ReportDetailsModel
            {
                PONumber         = dt.Rows[0]["PO_Number"].ToString()!,
                RateSequence     = dt.Rows[0]["RateSequence"].ToString()!,
                VendorCode       = dt.Rows[0]["VendorCode"].ToString()!,
                VendorName       = dt.Rows[0]["VendorName"].ToString()!,
                NoOfUnits        = dt.Rows[0]["NoOfUnits"].ToString()!,
                POEffectFrom     = dt.Rows[0]["PO_EffectFrom"].ToString()!,
                POEffectTo       = dt.Rows[0]["PO_EffectTo"].ToString()!,
                TotalMonthDays   = dt.Rows[0]["TotalMonthDays"].ToString()!,
                PayableDays      = dt.Rows[0]["PayableDays"].ToString()!,
                TotalWorkingDays = dt.Rows[0]["TotalWorkingDays"].ToString()!
            };
        }

        private void BuildTotalModel(DataTable dt, MonthlyPaymentReportViewModel vm, ReportSearchModel model)
        {
            if (dt.Rows.Count == 0) return;
            vm.TotalResourceCredit  = dt.Rows[0]["TotalResourceCredit"].ToString()  == "" ? "0.00" : dt.Rows[0]["TotalResourceCredit"].ToString()!;
            vm.TotalResourceDebit   = dt.Rows[0]["TotalResourceDebit"].ToString()   == "" ? "0.00" : dt.Rows[0]["TotalResourceDebit"].ToString()!;
            vm.TotalCompanyDebit    = dt.Rows[0]["TotalCompanyDebit"].ToString()    == "" ? "0.00" : dt.Rows[0]["TotalCompanyDebit"].ToString()!;
            vm.GrandPayable         = dt.Rows[0]["NetValue"].ToString()             == "" ? "0.00" : dt.Rows[0]["NetValue"].ToString()!;
            vm.DebitReason          = dt.Rows[0]["f_Reason"] != null ? dt.Rows[0]["f_Reason"].ToString()! : "";
            vm.IsGenerated          = dt.Rows[0]["f_IsGenerated"] is bool b && b;
        }

        private void InvoiceViewCount(string viewtype, string filename, ReportSearchModel model)
        {
            try
            {
                string Err = string.Empty;
                UserLogin? objUserLogin = HttpContext.Session.Get<UserLogin>("UserLogin");
                if (objUserLogin == null) return;

                var objPOProcessBAL = new POProcessBAL();
                var objPOProcessBO  = new POProcessBO
                {
                    VendorCode       = model.SelectedVendorCode,
                    VendorName       = model.SelectedVendorText,
                    PONo             = model.SelectedPOText,
                    RateCode         = model.SelectedPOValue,
                    MonthName        = model.SelectedMonth,
                    CreatedType      = objUserLogin.EntityType,
                    CreatedID        = objUserLogin.EntityID,
                    CreatedBy        = objUserLogin.CreatedBy,
                    CreatedByUName   = objUserLogin.CreatedByUname,
                    CreatedIP        = HttpContext.Connection.RemoteIpAddress?.ToString() ?? "",
                    SessionID        = Convert.ToInt32(objUserLogin.SessionId),
                    FormCode         = objUserLogin.FormCode,
                    FormName         = objUserLogin.FormName,
                    ViewType         = viewtype,
                    FileName         = filename
                };
                objPOProcessBAL.UPDATE_INVOICE_VIEW_LOG(objPOProcessBO, ref Err);
            }
            catch (Exception ex)
            {
                ExceptionUtility.LogException(ex, ex.Source ?? nameof(InvoiceViewCount));
            }
        }

        // ── PDF: iText 7 — Ledger orientation (replaces iTextSharp HTMLWorker) ──
        private static byte[] ConvertHtmlToPdfLedger(string htmlText, string savePath)
        {
            using var ms = new MemoryStream();
            using (var writer = new PdfWriter(ms))
            using (var pdf   = new PdfDocument(writer))
            {
                pdf.SetDefaultPageSize(PageSize.LEDGER.Rotate());
                HtmlConverter.ConvertToPdf(htmlText, pdf, new ConverterProperties());
            }
            var bytes = ms.ToArray();
            File.WriteAllBytes(savePath, bytes);
            return bytes;
        }

        // ── PDF: iText 7 — A4 portrait (covering letter) ─────────────────────
        private static byte[] ConvertHtmlToPdfA4(string htmlText, string savePath)
        {
            using var ms = new MemoryStream();
            using (var writer = new PdfWriter(ms))
            using (var pdf   = new PdfDocument(writer))
            {
                pdf.SetDefaultPageSize(PageSize.A4);
                HtmlConverter.ConvertToPdf(htmlText, pdf, new ConverterProperties());
            }
            var bytes = ms.ToArray();
            File.WriteAllBytes(savePath, bytes);
            return bytes;
        }

        // ── HTML builders — copied verbatim from code-behind ─────────────────

        private string GenerateStringForPO(DataTable dtDetails, DataTable dtGrid, DataTable dtTotal, ReportSearchModel model)
        {
            string strMainReport;
            string strpath = "../Images/MahindraLogo.bmp";

            string strTitle = $"<table style=\"width:100%;\"><Br><tr><td><img src=\"{strpath}\" style=\"width:60px;height:60px;\"></img></td></tr>"
                + $"<tr><td><p align=\"center\" style=\"font-size: 16; font-family: verdana; font-weight: bold; font-style: normal\">"
                + $"{dtDetails.Rows[0]["VendorName"]} (Vendor Code :{dtDetails.Rows[0]["LegacyCode"]})</p></td></tr></table>";

            string strAnnex = "<table style=\"width:100%;\"><tr><td><p align=\"center\" style=\"font-size: 16; font-family: verdana; font-weight: bold; font-style: normal\">Annex 3</p></td></tr></table>";

            string strPODetail = "<table style=\"width:100%;font-size: 10pt; font-family: verdana; font-weight: normal; font-style: normal; \"><tr>"
                + "<td style=\"width:60%;\"><table width=\"60%\" border=\"0\" cellspacing=\"0\"><tr><td style=\"width:60%;\">"
                + "<table style=\"width:60%;\" border=\"1\">"
                + $"<tr><td style=\"width:20%;font-weight: bold;\">PO Number / Rate Revision </td><td style=\"width:40%;\" align=\"left\"> ({model.SelectedPOText}) / {model.SelectedPORateSeq}</td></tr>"
                + $"<tr><td style=\"width:20%;font-weight: bold;\">Attendance Payable for the Month</td><td style=\"width:40%;\" align=\"left\"> {model.SelectedMonth}</td></tr>"
                + $"<tr><td style=\"width:20%;font-weight: bold;\">PO Effect From - To</td><td style=\"width:40%;\" align=\"left\">{dtDetails.Rows[0]["PO_EffectFrom"]} - {dtDetails.Rows[0]["PO_EffectTo"]}</td></tr>"
                + $"<tr><td style=\"width:20%;font-weight: bold;\">Resource Allowed || PO Payable Days</td><td style=\"width:40%;\" align=\"left\">{dtDetails.Rows[0]["ResourceCount"]} || {dtDetails.Rows[0]["PayableDays"]}</td></tr>"
                + "</table></td></tr></table></td>"
                + $"<td style=\"width: 40%\" align=\"right\" valign=\"top\"> <table width=\"40%\" border=\"1\" align=\"right\"><tr valign=\"top\"><td style=\"width: 40%;height: 28px;\" align=\"right\">"
                + $"<table border=\"0\" align=\"right\"><tr valign=\"top\"><td style=\"white-space: nowrap;\" align=\"right\">Printed Date</td><td>:</td><td>{DateTime.Now:dd-MMM-yyyy}</td></tr></table>"
                + "</td></tr></table></td></tr></table>";

            string strTotalDays = "<table style=\"width:100%;font-size: 10pt; font-family: verdana; font-weight: normal; font-style: normal; \"><tr><td style=\"width:60%;\">"
                + "<table width=\"60%\" border=\"0\" cellspacing=\"0\"><tr><td style=\"width:60%;\">"
                + "<table style=\"width:60%;\" border=\"1\">"
                + $"<tr><td style=\"white-space: nowrap;\">Total Month || H || WOFF || WRK Days</td><td style=\"white-space: nowrap;\" align=\"left\"> "
                + $"{dtDetails.Rows[0]["TotalMonthDays"]} || {dtDetails.Rows[0]["TotalHoliday"]} || {dtDetails.Rows[0]["TotalWeeklyOff"]} || {dtDetails.Rows[0]["TotalWorkingDays"]}</td></tr>"
                + "</table></td></tr></table></td>"
                + "<td style=\"width: 40%\" align=\"right\" valign=\"top\"><table width=\"40%\" border=\"0\" align=\"right\"><tr valign=\"top\">"
                + "<td style=\"width: 40%;height: 28px;\" align=\"right\"><table border=\"0\" align=\"right\">"
                + "<tr valign=\"top\"><td style=\"white-space: nowrap;\"></td><td></td><td></td></tr>"
                + "</table></td></tr></table></td></tr></table>";

            string strActualRecords = "";
            string strFooter = "";
            if (dtGrid != null && dtGrid.Rows.Count > 0)
            {
                strActualRecords = "<table style=\"width:100%;font-size: 9pt; font-family: verdana; \" border=\"1\" cellspacing=\"0\" cellpadding=\"0\">"
                    + "<tr><td style=\"width: 30%;white-space: nowrap;\">Resource Name</td><td style=\"width: 5%;\">Grade</td>"
                    + "<td style=\"width: 5%;\">Monthly PO Amount</td><td style=\"width: 5%;\">Maximum Fixed Amount</td>"
                    + "<td style=\"width: 5%;\">Maximum Variable %</td><td style=\"width: 5%;\">Approved Variable %</td>"
                    + "<td style=\"width: 5%;\">Approved Variable Amount</td><td style=\"width: 5%;\">Payable Days</td>"
                    + "<td style=\"width: 5%;\">Present Working Days</td><td style=\"width: 5%;\">Net Paid Days</td>"
                    + "<td style=\"width: 5%;\">Per Day Cost</td><td style=\"width: 5%;\">Total Amount</td>"
                    + "<td style=\"width: 5%;\">Deduction</td><td style=\"width: 5%;\">Net Payable</td></tr>";

                for (int i = 0; i < dtGrid.Rows.Count; i++)
                {
                    strActualRecords += "<tr>"
                        + $"<td style=\"width: 30%;white-space: nowrap;\">{dtGrid.Rows[i]["f_ResourceName"]}</td>"
                        + $"<td style=\"width: 5%;\" align=\"center\">{dtGrid.Rows[i]["f_GradeCode"]}</td>"
                        + $"<td style=\"width: 5%;\" align=\"center\">{dtGrid.Rows[i]["f_GradeMonthlyPOValue"]}</td>"
                        + $"<td style=\"width: 5%;\" align=\"center\">{dtGrid.Rows[i]["f_MaxFixedAmount"]}</td>"
                        + $"<td style=\"width: 5%;\" align=\"center\">{dtGrid.Rows[i]["f_MaxVariablePercentage"]}</td>"
                        + $"<td style=\"width: 5%; font-weight:bold;\" align=\"center\">{string.Format("{0:0.##}", dtGrid.Rows[i]["f_ActualVariablePercentagePDF"])}</td>"
                        + $"<td style=\"width: 5%;\" align=\"center\">{dtGrid.Rows[i]["f_ActualVariableAmount"]}</td>"
                        + $"<td style=\"width: 5%;\" align=\"center\">{dtGrid.Rows[i]["f_PayableDays"]}</td>"
                        + $"<td style=\"width: 5%;\" align=\"center\">{dtGrid.Rows[i]["f_PresentWorkingDays"]}</td>"
                        + $"<td style=\"width: 5%;\" align=\"center\">{dtGrid.Rows[i]["f_PaidDays"]}</td>"
                        + $"<td style=\"width: 5%; font-weight:bold;\" align=\"center\">{dtGrid.Rows[i]["f_PerDayCost"]}</td>"
                        + $"<td style=\"width: 5%; font-weight:bold;\" align=\"center\">{dtGrid.Rows[i]["f_CreditAmount"]}</td>"
                        + $"<td style=\"width: 5%; font-weight:bold;\" align=\"center\">{dtGrid.Rows[i]["f_DebitResource"]}</td>"
                        + $"<td style=\"width: 5%;font-weight:bold;\" align=\"center\">{dtGrid.Rows[i]["f_NetPayable"]}</td>"
                        + "</tr>";
                }
                strActualRecords += "</table>";

                string strFooterLeft = "<table width=\"60%\" border=\"1\" cellspacing=\"0\"><tr><td style=\"width:60%;\">"
                    + "<table style=\"width:60%;\" border=\"0\">"
                    + $"<tr><td style=\"white-space: nowrap;\">Total Resource Credit</td><td>:</td><td style=\"white-space: nowrap;\" align=\"left\"> {dtTotal.Rows[0]["TotalResourceCredit"]}</td></tr>"
                    + $"<tr><td style=\"white-space: nowrap;\">Total Resource Debit</td><td>:</td><td style=\"white-space: nowrap;\" align=\"left\"> {dtTotal.Rows[0]["TotalResourceDebit"]}</td></tr>"
                    + $"<tr><td style=\"white-space: nowrap;\">Total Company Debit</td><td>:</td><td style=\"white-space: nowrap;\" align=\"left\">{dtTotal.Rows[0]["TotalCompanyDebit"]}</td></tr>"
                    + $"<tr><td style=\"white-space: nowrap;\">Net Payable</td><td>:</td><td style=\"white-space: nowrap;\" align=\"left\">{dtTotal.Rows[0]["NetValue"]}</td></tr>"
                    + "</table></td></tr></table>";

                strFooter = "<table style=\"width:100%;font-size: 10pt; font-family: verdana; font-weight: normal; font-style: normal; \"><tr>"
                    + $"<td style=\"width:60%;\" valign=\"top\">{strFooterLeft} </td>"
                    + "<td style=\"width: 40%\" align=\"right\" valign=\"top\"></td>"
                    + "</tr></table>";
            }

            strMainReport = strTitle + "<br/>" + strAnnex + "<br/>" + strPODetail + "<br/>" + strTotalDays + "<br/>" + strActualRecords + "<br/>" + strFooter;
            strMainReport = "<div id=\"div1\" style=\"width:100%\">" + strMainReport + "</div>";
            return strMainReport;
        }

        private string GenerateCoverLetter(DataTable dtDetails, DataTable dtGrid, DataTable dtTotal, DataTable dtPOTax, ReportSearchModel model)
        {
            string[] selectedMonth = model.SelectedMonth.Trim().Split(',');
            int mnth = DateTime.ParseExact(selectedMonth[0].Trim(), "MMMM", CultureInfo.CurrentCulture).Month;
            int yr   = Convert.ToInt32(selectedMonth[1].Trim());
            string mnthname     = GetMonthName(mnth, yr);
            DateTime dtAttFrom  = GetFirstDayOfMonth(mnth, yr);
            DateTime dtAttTo    = GetLastDayOfMonth(mnth, yr);
            string firstdaydate = $"{dtAttFrom.Day} {mnthname},{yr}";
            string lastdaydate  = $"{dtAttTo.Day} {mnthname},{yr}";

            if (dtPOTax.Rows.Count == 0) return "";

            string strAnnex = "<table style=\"width:100%;\"><tr><td><p align=\"center\" style=\"font-size: 10; font-family: verdana; font-weight: bold; font-style: normal\">Annex 1</p></td></tr></table>";

            string strCoverLetter = "<table style=\"width:100%;\">"
                + $"<tr><td><p align=\"center\" style=\"font-size: 16; font-family: verdana; font-weight: normal; font-style: normal\"> {dtDetails.Rows[0]["VendorName"]}</p></td></tr>"
                + $"<tr><td><p align=\"center\" style=\"font-size: 16; font-family: verdana; font-weight: normal; font-style: normal\"> ( Vendor Code : {dtDetails.Rows[0]["LegacyCode"]} ) </p></td></tr>"
                + "</table>";

            string strCoverDate = "<table style=\"width:100%;\">"
                + $"<tr><td><p align=\"left\" style=\"font-size: 10; font-family: verdana; font-weight: normal; font-style: normal\"> Dated : {GetFirstDayOfMonth(DateTime.Now.Month, DateTime.Now.Year):dd/MMM/yyyy}</p></td></tr>"
                + "</table>";

            string strCoverTo = "<table style=\"width:100%;\">"
                + "<tr><td><p align=\"left\" style=\"font-size: 10; font-family: verdana; font-weight: normal; font-style: normal\"> To : The Department of BITS\\Purchase\\Accounts\\Tax</p></td></tr>"
                + "</table>";

            string strCoverSub = "<table style=\"width:100%;\">"
                + $"<tr><td><p align=\"left\" style=\"font-size: 10; font-family: verdana; font-weight: normal; font-style: normal\"> Sub : Release of payments to {dtDetails.Rows[0]["VendorName"]}</p></td></tr>"
                + "</table>";

            string strCoverRespected = "<table style=\"width:100%;\">"
                + "<tr><td><p align=\"left\" style=\"font-size: 10; font-family: verdana; font-weight: normal; font-style: normal\"> Dear Sir, </p></td></tr>"
                + "</table>";

            string strCoverBody1 = $"<table style=\"width:100%;\"><tr><td><p align=\"left\" style=\"font-size: 10; font-family: verdana; font-weight: normal; font-style: normal\">"
                + $"<b>{dtDetails.Rows[0]["VendorName"]}</b> has been awarded with the order for providing technical resources in different <br/> "
                + $"category as per the purchase order number – {model.SelectedPOText} dated {dtDetails.Rows[0]["PO_Date"]} for the period <br/> "
                + $"{dtDetails.Rows[0]["PO_EffectFrom"]} to {dtDetails.Rows[0]["PO_EffectTo"]}. </p></td></tr></table>";

            string invoiceNum = dtPOTax.Rows[0]["f_InvoiceNumber"] != null && dtPOTax.Rows[0]["f_InvoiceNumber"].ToString() != ""
                ? dtPOTax.Rows[0]["f_InvoiceNumber"].ToString()!
                : "_______________";

            string strCoverBody2 = $"<table style=\"width:100%;\"><tr><td><p align=\"left\" style=\"font-size: 10; font-family: verdana; font-weight: normal; font-style: normal\">"
                + $" The original invoice ( {invoiceNum} ) from the vendor is attached in  <b>'Annex2'</b> from <b>'{dtDetails.Rows[0]["VendorName"]}'</b>"
                + $"  dated {dtPOTax.Rows[0]["f_InvoiceDate"]}. </p></td></tr></table>";

            string strCoverBody3 = $"<table style=\"width:100%;\"><tr><td><p align=\"left\" style=\"font-size: 10; font-family: verdana; font-weight: normal; font-style: normal\">"
                + $" This letter is to authorize that the resources have been coming to MMFSL for the period of <br/>"
                + $"{firstdaydate} to {lastdaydate}, details as attached in <b>'Annex3'</b>.</p></td></tr></table>";

            string strCoverBody4 = $"<table style=\"width:100%;\"><tr><td><p align=\"left\" style=\"font-size: 10; font-family: verdana; font-weight: normal; font-style: normal\">"
                + $" We request you to please release the payments for the below given invoices to <b>'{dtDetails.Rows[0]["VendorName"]}'</b>.</p></td></tr></table>";

            string strCoverPayment = "<table style=\"width:100%;font-size: 10pt; font-family: verdana; font-weight: normal; font-style: normal; \"><tr>"
                + "<td style=\"width:80%;\"><table width=\"80%\" border=\"0\" cellspacing=\"0\"><tr><td style=\"width:80%;\">"
                + "<table style=\"width:100%;\" border=\"1\">"
                + $"<tr><td style=\"white-space: nowrap;\">Total Amount</td><td style=\"white-space: nowrap;\" align=\"left\">{dtPOTax.Rows[0]["TotalAmount"]}</td></tr>"
                + $"<tr><td style=\"white-space: nowrap;\">Tax Percentage</td><td style=\"white-space: nowrap;\" align=\"left\">{dtPOTax.Rows[0]["TaxPercentage"]}</td></tr>"
                + $"<tr><td style=\"white-space: nowrap;\">Tax Amount</td><td style=\"white-space: nowrap;\" align=\"left\">{dtPOTax.Rows[0]["TAxAmount"]}</td></tr>"
                + $"<tr><td style=\"white-space: nowrap;\">Tax On Gross</td><td style=\"white-space: nowrap;\" align=\"left\">{dtPOTax.Rows[0]["TaxOnGrossAmount"]}</td></tr>"
                + $"<tr><td nowrap=\"nowrap\"> Net Amount to be Payable</td><td style=\"white-space: nowrap;\" align=\"left\">{dtPOTax.Rows[0]["NetPayableAmount"]}</td></tr>"
                + "</table></td></tr></table></td>"
                + "<td style=\"width: 20%\" align=\"right\" valign=\"top\"></td></tr></table>";

            return strCoverLetter + "<br />" + strAnnex + "<br/>" + strCoverDate + "<br/>" + strCoverTo
                + "<br/>" + strCoverSub + "<br/>" + strCoverRespected + "<br/>" + strCoverBody1
                + "<br/>" + strCoverBody2 + "<br/>" + strCoverBody3 + "<br/>" + strCoverBody4
                + "<br/>" + strCoverPayment;
        }

        // ── Date helpers — copied verbatim ────────────────────────────────────
        private static string GetMonthName(int month, int year) =>
            new DateTime(year, month, 1).ToString("MMMM");

        private static DateTime GetFirstDayOfMonth(int iMonth, int iYear)
        {
            DateTime dtFrom = new DateTime(iYear, iMonth, 1);
            dtFrom = dtFrom.AddDays(-(dtFrom.Day - 1));
            return dtFrom;
        }

        private static DateTime GetLastDayOfMonth(int iMonth, int iYear)
        {
            DateTime dtTo = new DateTime(iYear, iMonth, 1).AddMonths(1).AddDays(-1);
            return dtTo;
        }
    }
}
