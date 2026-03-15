using Microsoft.AspNetCore.Mvc.Rendering;
using System.Collections.Generic;
using System.Data;

namespace SYS_VENDOR_MGMT.Models
{
    /// <summary>
    /// Drives the GET view — dropdown population and results display.
    /// </summary>
    public class MonthlyPaymentReportViewModel
    {
        public List<SelectListItem> Months { get; set; } = new();

        // Search context echoed back after POST
        public ReportSearchModel Search { get; set; } = new();

        // Panel visibility
        public bool ShowDetails     { get; set; }
        public bool IsGenerated     { get; set; }
        public bool ShowCoveringBtn { get; set; }

        // Header / error
        public string Header       { get; set; } = "";
        public string ErrorMessage { get; set; } = "";

        // Detail labels (mirror original ASPX Label controls)
        public ReportDetailsModel Details { get; set; } = new();

        // Grid data
        public DataTable? GridRows { get; set; }

        // Total summary
        public string TotalResourceCredit { get; set; } = "0.00";
        public string TotalResourceDebit  { get; set; } = "0.00";
        public string TotalCompanyDebit   { get; set; } = "0.00";
        public string GrandPayable        { get; set; } = "0.00";
        public string DebitReason         { get; set; } = "";
    }

    /// <summary>
    /// Mirrors the original ASPX Label fields shown in pnlDetails.
    /// </summary>
    public class ReportDetailsModel
    {
        public string PONumber         { get; set; } = "";
        public string RateSequence     { get; set; } = "";
        public string VendorCode       { get; set; } = "";
        public string VendorName       { get; set; } = "";
        public string NoOfUnits        { get; set; } = "";
        public string POEffectFrom     { get; set; } = "";
        public string POEffectTo       { get; set; } = "";
        public string TotalMonthDays   { get; set; } = "";
        public string PayableDays      { get; set; } = "";
        public string TotalWorkingDays { get; set; } = "";
    }

    /// <summary>
    /// Form values posted from Search / Generate / Export buttons.
    /// Replaces the ViewState-preserved dropdown selections.
    /// </summary>
    public class ReportSearchModel
    {
        public string SelectedMonth       { get; set; } = "";
        public string SelectedVendorCode  { get; set; } = "";
        public string SelectedVendorText  { get; set; } = "";
        public string SelectedPOValue     { get; set; } = "";  // f_RateCode
        public string SelectedPOText      { get; set; } = "";  // f_PONumber
        public string SelectedPORateSeq   { get; set; } = "";  // displayed in PDF
        public string DebitReason         { get; set; } = "";
    }
}
