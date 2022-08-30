using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace AssetTrackingApplication
{
    public partial class AssetTrackingForm : Form {
        // set file name
        const string _fullName = @"C:\AssetTracking\AssetTrackingTable.xlsx";
        Excel _excel = new Excel(_fullName);
        public AssetTrackingForm() {
            InitializeComponent();
        }

        #region insertData_buttonEvents
        private void btn_InitializeAssetInsertion_Click(object sender, EventArgs e) {
            ToggleInsertionControls(true);
            ToggleMainControls(false);
            ClearInsertionDataTextBoxes();
            ClearUpdateDataTextBoxes();
            
            var assetName = cb_assetName.Text;
            var assetList = GetAssetList();
            var columns = GetColumnList();
            var assetRow = GetAssetRow(assetList, assetName);
            InsertPreviousInsertionValuesToTextboxes(_excel.Worksheet, assetRow, columns);

        }

        private void btn_insertData_Click(object sender, EventArgs e) {
            // deactivate controls
            ToggleInsertionControls(false);
            var assetList = GetAssetList();
            var columns = GetColumnList();
            try {
                // Get parameters for storing asset in _excel (parameters from excel, winforms)
                var assetName = cb_assetName.Text;
                var assetRow = GetAssetRow(assetList, assetName);
                var assetClass = txt_assetClass.Text;
                var initialAmount = GetInitialAmount(_excel.Worksheet, assetRow, columns);
                var currentAmount = GetCurrentAmount();
                var previousSharePrice = GetSharePriceForInsertion(_excel.Worksheet, assetRow, columns);
                var previousCostBasis = GetPreviousCostBasis(_excel.Worksheet, assetRow, columns);
                var costBasisParameter = GetCostBasisParameter();
                if (costBasisParameter == null)
                    return;
                // Create instance of AssetInvestment to insert 
                var assetInvestment = new AssetInvestment(assetName, assetClass, initialAmount, currentAmount, costBasisParameter.InsertionPrice, 
                                                          previousCostBasis, costBasisParameter.InvestedCapital, previousSharePrice);
                // insert values into _excel list
                assetInvestment.InsertAssetInvestment(assetInvestment, assetRow, _excel.Worksheet, columns, assetList);
                WriteInsertionDataToTextBox(assetInvestment);
            } finally {
                ToggleMainControls(true);
                ToggleInsertionControls(false);
            }
        }
        #endregion insertData_buttonEvents

        #region updateData_buttonEvents
        private void btn_InitializeUpdate_Click(object sender, EventArgs e) {
            ToggleUpdateControls(true);
            ToggleMainControls(false);
            ClearUpdateDataTextBoxes();
            ClearInsertionDataTextBoxes();

            var assetName = cb_assetName.Text;
            var assetList = GetAssetList();
            var columns = GetColumnList();
            var assetRow = GetAssetRow(assetList, assetName);
            InsertPreviousUpdateValuesToTextboxes(_excel.Worksheet, assetRow, columns);
        }
        private void btn_updateData_Click(object sender, EventArgs e) {
            ToggleUpdateControls(false);
            var assetList = GetAssetList();
            var columns = GetColumnList();

            try {
                // Get parameters for storing asset in _excel (parameters from _excel, winforms)
                var assetName = cb_assetName.Text;
                var assetRow = GetAssetRow(assetList, assetName);
                var assetClass = txt_assetClass.Text;
                var amount = GetInitialAmount(_excel.Worksheet, assetRow, columns);
                var investedCapital = GetInvestedCapital(_excel.Worksheet, assetRow, columns);
                var previousValue = GetPreviousValue(_excel.Worksheet, assetRow, columns);
                var currentSharePrice = GetCurrentSharePrice();
                // Create instance of AssetUpdate
                var assetUpdate = new AssetUpdate(assetName, assetClass, currentSharePrice, amount, previousValue, investedCapital);
                // insert values into _excel list and winforms
                WriteUpdateDataToTextBox(assetUpdate);
                assetUpdate.InsertAssetUpdate(assetUpdate, assetRow, _excel.Worksheet, columns, assetList);
            }
            finally {
                ToggleMainControls(true);
                ToggleUpdateControls(false);
            }
        }
        #endregion updateData_buttonEvents

        #region toggleControls
        public void ToggleInsertionControls(bool state) {
            txt_currentAmount.Enabled = state;
            btn_insertData.Enabled = state;
            txt_insertionPrice.Enabled = state;
            txt_investedCapital.Enabled = state;
        }

        public void ToggleMainControls(bool state) {
            cb_assetName.Enabled = state;
            txt_assetClass.Enabled = state;
            btn_InitializeUpdate.Enabled = state;
            btn_InitializeAssetInsertion.Enabled = state;
        }

        public void ToggleUpdateControls(bool state) {
            txt_currentSharePrice.Enabled = state;
            btn_updateData.Enabled = state;
        }
        #endregion toggleControls

        #region File management for excel data
        // AssetList - insert new Assets here:
        public Dictionary<string, int> GetAssetList() {
            var fileContent = File.ReadAllText("AssetList.json");
            var assetList = JsonConvert.DeserializeObject<Dictionary<string, int>>(fileContent);

            return assetList;
        }

        public Dictionary<string, string> GetColumnList() {
            var fileContent = File.ReadAllText("ColumnList.json");

            var columns = JsonConvert.DeserializeObject<Dictionary<string, string>>(fileContent);

            return columns;
        }
        #endregion File management for excel data

        #region updateTextBoxValues
        public void WriteUpdateDataToTextBox(AssetUpdate assetUpdate) {
            txt_previousPrice.Text = assetUpdate.PreviousValue.ToString();
            txt_currentValue.Text = assetUpdate.TotalValue.ToString();
            txt_priceGain.Text = assetUpdate.GainTotal.ToString();
            txt_percentalGain.Text = (assetUpdate.GainRelative * 100).ToString();
        }
        public void ClearUpdateDataTextBoxes() {
            txt_previousPrice.Text = "";
            txt_currentValue.Text = "";
            txt_currentSharePrice.Text = "";
            txt_priceGain.Text = "";
            txt_percentalGain.Text = "";
        }
        
        public void WriteInsertionDataToTextBox(AssetInvestment assetInvestment) {
            txt_previousAmount.Text = assetInvestment.InitialAmount.ToString();
            txt_investedCapital.Text = assetInvestment.InvestedCapital.ToString();
        }
        public void ClearInsertionDataTextBoxes() {
            txt_currentAmount.Text = "";
            txt_previousAmount.Text = "";
            txt_insertionPrice.Text = "";
            txt_investedCapital.Text = "";
        }

        public void InsertPreviousInsertionValuesToTextboxes(Worksheet _worksheet, int assetRow, Dictionary<string, string> columns) {
            
            var previousAmount = (Range)_worksheet.Cells[assetRow, columns["Amount"]];
            if (previousAmount.Value != null)
                txt_previousAmount.Text = previousAmount.Value.ToString();
            else
                txt_previousAmount.Text = "";
        }
        public void InsertPreviousUpdateValuesToTextboxes(Worksheet _worksheet, int assetRow, Dictionary<string, string> columns) {
            var previousValue = (Range)_worksheet.Cells[assetRow, columns["TotalValue"]];
            if (previousValue.Value != null)
                txt_previousPrice.Text = previousValue.Value.ToString();
            else
                txt_previousSharePrice.Text = "";

            var previousSharePrice = (Range)_worksheet.Cells[assetRow, columns["PricePerShare"]];
            if (previousSharePrice.Value != null)
                txt_previousSharePrice.Text = previousSharePrice.Value.ToString();
            else
                txt_previousSharePrice.Text = "";
        }

        #endregion updateTextBoxValues

        #region retrieveExcelCellValues
        public decimal GetInitialAmount(Worksheet excel, int assetRow, Dictionary<string, string> columns) {
            var excelCell = (Range)excel.Cells[assetRow, columns["Amount"]];
            if (excelCell.Value != null)
                return (decimal) excelCell.Value;
            else
                return 0;
        }
        public decimal GetInvestedCapital(Worksheet excel, int assetRow, Dictionary<string, string> columns) {
            var excelCell = (Range)excel.Cells[assetRow, columns["InvestedCapital"]];
            if (excelCell.Value != null)
                return (decimal)excelCell.Value;
            else
                return 0;
        }
        public decimal GetPreviousValue(Worksheet excel, int assetRow, Dictionary<string, string> columns) {
            var valueCell = (Range) excel.Cells[assetRow, columns["TotalValue"]];
            if (valueCell.Value != null)
                return (decimal) valueCell.Value;
            else
                return 0;
        }
        public decimal GetSharePriceForInsertion(Worksheet excel, int assetRow, Dictionary<string, string> columns) {
            var valueCell = (Range)excel.Cells[assetRow, columns["PricePerShare"]];
            if (valueCell.Value != null)
                return (decimal) valueCell.Value;
            else
                return 1;
        }

        public decimal GetPreviousCostBasis(Worksheet excel, int assetRow, Dictionary<string, string> columns) {
            var valueCell = (Range) excel.Cells[assetRow, columns["CostBasis"]];
            if (valueCell.Value != null)
                return (decimal) valueCell.Value;
            else
                return 0;
        }

        public decimal GetTotalValue(Worksheet excel, Dictionary<string, string> columns)
        {
            var valueCell = (Range)excel.Cells[150, columns["TotalValue"]];
            if (valueCell.Value != null)
                return (decimal)valueCell.Value;
            else
                return 0;
        }

        public CostBasisParameter GetCostBasisParameter() {
            if (txt_insertionPrice.Text == "" && txt_investedCapital.Text == "") {
                MessageBox.Show("Please enter either invested capital or insertion price for the asset!");
                return null;
            } 
            if (txt_investedCapital.Text == "") {
                return new CostBasisParameter(Convert.ToDecimal(txt_insertionPrice.Text), 0);
            }
            return new CostBasisParameter(0, Convert.ToDecimal(txt_investedCapital.Text));
        }
        #endregion retrieveExcelCellValues

        #region retrieveWinFormsValues
        public decimal GetCurrentAmount() {
            decimal currentAmount;
            try {
                currentAmount = Convert.ToDecimal(txt_currentAmount.Text);
            }
            catch (Exception exception) {
                Console.WriteLine("Please enter a valid number next time! Exception: " + exception.Message);
                throw;
            }
            return currentAmount;
        }

        public decimal GetCurrentSharePrice() {
            decimal sharePrice;
            try {
                sharePrice = Convert.ToDecimal(txt_currentSharePrice.Text);
            }
            catch (Exception exception) {
                MessageBox.Show("Please enter a valid capital next time! Exception: " + exception.Message);
                throw;
            }
            return sharePrice;
        }
        public decimal GetInsertionPrice() {
            decimal insertionPrice;
            try {
                insertionPrice = Convert.ToDecimal(txt_insertionPrice.Text);
            }
            catch (Exception exception) {
                MessageBox.Show("Please enter a valid capital next time! Exception: " + exception.Message.ToString());
                throw;
            }
            return insertionPrice;
        }

        #endregion retrieveWinFormsValues

        #region excelFileManagement

        public void DeactivateControls() {
            ToggleMainControls(false);
            ToggleInsertionControls(false);
            ToggleUpdateControls(false);

        }
       
        #endregion

        #region assetCharacteristics
        public int GetAssetRow(Dictionary<string, int> assetList, string assetName) {
            if (!assetList.TryGetValue(assetName, out var assetRow)) {
                MessageBox.Show("No asset '" + assetName + "' defined - please add it to the dictionary!");
            }
            return assetRow;
        }
        public string GetAssetClass(int assetRow, string assetName) {
            var assetClasses = AssetClass.GetAllAssetClasses();
            var assetClass = assetClasses.Find(a => (a.FirstRow <= assetRow) && (a.LastRow >= assetRow)); 
            
            return assetClass.Name;
        }
        private void cb_assetName_MouseClick(object sender, MouseEventArgs e) {
            var assetList = GetAssetList();
            cb_assetName.DataSource = assetList.Keys.ToList();
        }

        private void cb_assetName_TextChanged(object sender, EventArgs e) {
            var assetList = GetAssetList();
            var assetName = cb_assetName.Text;
            var assetRow = GetAssetRow(assetList, assetName);
            var assetClass = GetAssetClass(assetRow, assetName);
            txt_assetClass.Text = assetClass;
        }
        #endregion assetCharacteristics

        private void btn_save_Click(object sender, EventArgs e)
        {
            txt_totalValue.Text = GetTotalValue(_excel.Worksheet, GetColumnList()).ToString();
            DeactivateControls();
            _excel.CloseExcelFile();
            btn_save.Enabled = false;
        }

        private void btn_createAssetClass_Click(object sender, EventArgs e)
        {
            var assetClassForm = new AssetClassForm(AssetClass.GetAllAssetClasses());
            assetClassForm.ShowDialog();
        }
    }
}
