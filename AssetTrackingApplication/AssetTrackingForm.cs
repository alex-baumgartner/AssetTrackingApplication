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
using Microsoft.WindowsAPICodePack.Dialogs;
using Newtonsoft.Json;

namespace AssetTrackingApplication
{
    public partial class AssetTrackingForm : Form {
        Excel _excel; 
        List<AssetClass> _assetClasses;
        Dictionary<string, int> _assets;
        Dictionary<string, string> _columns;
        string _directoryPath;
        public AssetTrackingForm() {
            InitializeComponent();
            var excelFileName = GetExcelFilePathFromUser();
            _directoryPath = GetDirectoryPathFromUser();
            _excel = new Excel(excelFileName);
            _assetClasses = AssetClass.GetAllAssetClasses(_directoryPath);
            _assets = AssetClass.GetAssetList(_directoryPath);
            _columns = GetColumnList(_directoryPath);
        }

        private string GetDirectoryPathFromUser()
        {
            var applicationPath = System.Windows.Forms.Application.StartupPath;

            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.Title = "Choose data folder";
            dialog.InitialDirectory = applicationPath;
            dialog.IsFolderPicker = true;
            try
            {
                dialog.ShowDialog();
            }
            catch (DirectoryNotFoundException ex)
            {
                MessageBox.Show("Could not find folder:  " + dialog.FileName + ",  " + ex.Message, null, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurred when choosing data folder, " + ex.Message, null, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return dialog.FileName;
        }
        #region insertData_buttonEvents
        private void btn_InitializeAssetInsertion_Click(object sender, EventArgs e) {
            ToggleInsertionControls(true);
            ToggleMainControls(false);
            ClearInsertionDataTextBoxes();
            ClearUpdateDataTextBoxes();
            
            var assetName = cb_assetName.Text;
            
            var assetRow = GetAssetRow(assetName);
            InsertPreviousInsertionValuesToTextboxes(assetRow);

        }

        private void btn_insertData_Click(object sender, EventArgs e) {
            // deactivate controls
            ToggleInsertionControls(false);
            try {
                // Get parameters for storing asset in _excel (parameters from excel, winforms)
                var assetName = cb_assetName.Text;
                var assetRow = GetAssetRow(assetName);
                var assetClass = txt_assetClass.Text;
                var initialAmount = GetInitialAmount(assetRow);
                var currentAmount = GetCurrentAmount();
                var previousSharePrice = GetSharePriceForInsertion(assetRow);
                var previousCostBasis = GetPreviousCostBasis(assetRow);
                var costBasisParameter = GetCostBasisParameter();
                if (costBasisParameter == null)
                    return;
                // Create instance of AssetInvestment to insert 
                var assetInvestment = new AssetInvestment(assetName, assetClass, initialAmount, currentAmount, costBasisParameter.InsertionPrice, 
                                                          previousCostBasis, costBasisParameter.InvestedCapital, previousSharePrice);
                // insert values into _excel list
                assetInvestment.InsertAssetInvestment(assetInvestment, assetRow, _excel.Worksheet, _columns, _assets);
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
            var assetRow = GetAssetRow(assetName);
            InsertPreviousUpdateValuesToTextboxes(assetRow);
        }
        private void btn_updateData_Click(object sender, EventArgs e) {
            ToggleUpdateControls(false);

            try {
                // Get parameters for storing asset in _excel (parameters from _excel, winforms)
                var assetName = cb_assetName.Text;
                var assetRow = GetAssetRow(assetName);
                var assetClass = txt_assetClass.Text;
                var amount = GetInitialAmount(assetRow);
                var investedCapital = GetInvestedCapital(assetRow);
                var previousValue = GetPreviousValue(assetRow);
                var currentSharePrice = GetCurrentSharePrice();
                // Create instance of AssetUpdate
                var assetUpdate = new AssetUpdate(assetName, assetClass, currentSharePrice, amount, previousValue, investedCapital);
                // insert values into _excel list and winforms
                WriteUpdateDataToTextBox(assetUpdate);
                assetUpdate.InsertAssetUpdate(assetUpdate, assetRow, _excel.Worksheet, _columns, _assets);
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
        public Dictionary<string, string> GetColumnList(string path) {
            var fileContent = File.ReadAllText(path + "/ColumnList.json");

            var columns = JsonConvert.DeserializeObject<Dictionary<string, string>>(fileContent);

            return columns;
        }

        private string GetExcelFilePathFromUser()
        {
            string filePath;
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Choose Excel file";
                var applicationPath = System.Windows.Forms.Application.StartupPath;
                if (Directory.Exists(applicationPath))
                {
                    openFileDialog.InitialDirectory = applicationPath;
                }
                else
                {
                    openFileDialog.InitialDirectory = @"C:\";
                }
                openFileDialog.ShowDialog();
                filePath = openFileDialog.FileName;
            }
            return filePath;
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

        public void InsertPreviousInsertionValuesToTextboxes(int assetRow) {
            
            var previousAmount = (Range)_excel.Worksheet.Cells[assetRow, _columns["Amount"]];
            if (previousAmount.Value != null)
                txt_previousAmount.Text = previousAmount.Value.ToString();
            else
                txt_previousAmount.Text = "";
        }
        public void InsertPreviousUpdateValuesToTextboxes(int assetRow) {
            var previousValue = (Range)_excel.Worksheet.Cells[assetRow, _columns["TotalValue"]];
            if (previousValue.Value != null)
                txt_previousPrice.Text = previousValue.Value.ToString();
            else
                txt_previousSharePrice.Text = "";

            var previousSharePrice = (Range)_excel.Worksheet.Cells[assetRow, _columns["PricePerShare"]];
            if (previousSharePrice.Value != null)
                txt_previousSharePrice.Text = previousSharePrice.Value.ToString();
            else
                txt_previousSharePrice.Text = "";
        }

        #endregion updateTextBoxValues

        #region retrieveExcelCellValues
        public decimal GetInitialAmount(int assetRow) {
            var excelCell = (Range)_excel.Worksheet.Cells[assetRow, _columns["Amount"]];
            if (excelCell.Value != null)
                return (decimal) excelCell.Value;
            else
                return 0;
        }
        public decimal GetInvestedCapital(int assetRow) {
            var excelCell = (Range)_excel.Worksheet.Cells[assetRow, _columns["InvestedCapital"]];
            if (excelCell.Value != null)
                return (decimal)excelCell.Value;
            else
                return 0;
        }
        public decimal GetPreviousValue(int assetRow) {
            var valueCell = (Range)_excel.Worksheet.Cells[assetRow, _columns["TotalValue"]];
            if (valueCell.Value != null)
                return (decimal) valueCell.Value;
            else
                return 0;
        }
        public decimal GetSharePriceForInsertion(int assetRow) {
            var valueCell = (Range)_excel.Worksheet.Cells[assetRow, _columns["PricePerShare"]];
            if (valueCell.Value != null)
                return (decimal) valueCell.Value;
            else
                return 1;
        }

        public decimal GetPreviousCostBasis(int assetRow) {
            var valueCell = (Range)_excel.Worksheet.Cells[assetRow, _columns["CostBasis"]];
            if (valueCell.Value != null)
                return (decimal) valueCell.Value;
            else
                return 0;
        }

        public decimal GetTotalValue()
        {
            var valueCell = (Range)_excel.Worksheet.Cells[150, _columns["TotalValue"]];
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

        public void DeactivateControls() {
            ToggleMainControls(false);
            ToggleInsertionControls(false);
            ToggleUpdateControls(false);

        }

        #region assetCharacteristics
        public int GetAssetRow(string assetName) {
            if (!_assets.TryGetValue(assetName, out var assetRow)) {
                MessageBox.Show("No asset '" + assetName + "' defined - please add it to the dictionary!");
            }
            return assetRow;
        }
        public string GetAssetClass(int assetRow, string assetName) {
            var assetClasses = AssetClass.GetAllAssetClasses(_directoryPath);
            var assetClass = assetClasses.Find(a => (a.FirstRow <= assetRow) && (a.LastRow >= assetRow)); 
            
            return assetClass.Name;
        }
        private void cb_assetName_MouseClick(object sender, MouseEventArgs e) {
            cb_assetName.DataSource = _assets.Keys.ToList();
        }

        private void cb_assetName_TextChanged(object sender, EventArgs e) {
            var assetName = cb_assetName.Text;
            var assetRow = GetAssetRow(assetName);
            var assetClass = GetAssetClass(assetRow, assetName);
            txt_assetClass.Text = assetClass;
        }
        #endregion assetCharacteristics

        private void btn_finish_Click(object sender, EventArgs e)
        {
            txt_totalValue.Text = GetTotalValue().ToString();
            DeactivateControls();
            btn_finish.Enabled = false;
        }

        private void btn_createAssetClass_Click(object sender, EventArgs e)
        {
            var assetClassForm = new AssetClassForm(_assetClasses);
            assetClassForm.ShowDialog();
        }

        private void btn_createAsset_Click(object sender, EventArgs e)
        {
            var assetForm = new AssetForm(_assetClasses, _assets);
            assetForm.ShowDialog();
            if (assetForm.Assets != null)
            {
                _assets = assetForm.Assets;
            }
        }

        private void AssetTrackingForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (AssetClass.AssetContentChanged(_assets, _assetClasses, _directoryPath))
            {
                var saveChanges = MessageBox.Show("Do you want to save changed to assets and asset classes?", "Changes detected!", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (saveChanges == DialogResult.Yes)
                {
                    AssetClass.UpdateAssetClassFile(_assetClasses, _directoryPath);
                    AssetClass.UpdateAssetFile(_assets, _directoryPath);
                }
            }

            if (_excel != null)
            {
                _excel.CloseExcelFile();
            }
        }
    }
}
