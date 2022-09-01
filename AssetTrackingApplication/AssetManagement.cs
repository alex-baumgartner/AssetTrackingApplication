using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace AssetTrackingApplication
{
    public class AssetUpdate : Asset
    {
        public AssetUpdate(string name, string assetClass, decimal pricePerShare, decimal amount, decimal previousValue, decimal investedCapital) : base(name, assetClass) {
            Name = name;
            AssetClass = assetClass;
            PricePerShare = pricePerShare;
            Amount = amount;
            PreviousValue = previousValue;
            InvestedCapital = investedCapital;
        }

        public decimal PricePerShare { get; set; }
        public decimal Amount { get; set; }
        public decimal TotalValue => PricePerShare * Amount;
        public decimal PreviousValue { get; set; }
        public decimal GainTotal => TotalValue - PreviousValue;
        public decimal GainRelative => PreviousValue != 0 ? ((TotalValue / PreviousValue) - 1) : TotalValue == 0 ? 0 : 1;
        public decimal InvestedCapital { get; set; }
        public decimal Performance => InvestedCapital != 0 ? (TotalValue - InvestedCapital) / InvestedCapital :0;


        private void InsertPricePerShare(AssetUpdate assetUpdate, string assetColumn, int assetRow, Worksheet excel) {
            excel.Cells[assetRow, assetColumn] = assetUpdate.PricePerShare;
        }
        private void InsertTotalValue(AssetUpdate assetUpdate, string assetColumn, int assetRow, Worksheet excel) {
            excel.Cells[assetRow, assetColumn] = assetUpdate.TotalValue;
        }

        private void InsertPreviousValue(AssetUpdate assetUpdate, string assetColumn, int assetRow, Worksheet excel) {
            excel.Cells[assetRow, assetColumn] = assetUpdate.PreviousValue;
        }

        private void InsertGainTotal(AssetUpdate assetUpdate, string assetColumn, int assetRow, Worksheet excel) {
            excel.Cells[assetRow, assetColumn] = assetUpdate.GainTotal;
        }

        private void InsertGainRelative(AssetUpdate assetUpdate, string assetColumn, int assetRow, Worksheet excel) {
            excel.Cells[assetRow, assetColumn] = assetUpdate.GainRelative;
        }

        private void InsertPerformance(AssetUpdate assetUpdate, string assetColumn, int assetRow, Worksheet excel) {
            excel.Cells[assetRow, assetColumn] = assetUpdate.Performance;
        }
        public void InsertAssetUpdate(AssetUpdate assetUpdate, int assetRow, Worksheet excel, Dictionary<string, string> columns, Dictionary<string, int> assets) {
            InsertPricePerShare(assetUpdate, columns["PricePerShare"], assetRow, excel);
            InsertTotalValue(assetUpdate, columns["TotalValue"], assetRow, excel);
            InsertPreviousValue(assetUpdate, columns["PreviousValue"], assetRow, excel);
            InsertGainTotal(assetUpdate, columns["GainTotal"], assetRow, excel);
            InsertGainRelative(assetUpdate, columns["GainRelative"], assetRow, excel);
            InsertPerformance(assetUpdate, columns["Performance"], assetRow, excel);
            InsertAssetShare(excel, columns, assets);
        }

        public void InsertAssetShare(Worksheet excel, Dictionary<string, string> columns, Dictionary<string, int> assets) {
            const int lastRow = 150;
            foreach (var assetRow in assets.Values) {
                var assetValueCell = (Range)excel.Cells[assetRow, columns["TotalValue"]];
                decimal assetValue;
                if (assetValueCell.Value != null)
                    assetValue = (decimal)assetValueCell.Value;
                else
                    assetValue = 0;

                var totalValueCell = (Range)excel.Cells[lastRow, columns["TotalValue"]];
                var totalValue = (decimal)totalValueCell.Value;

                excel.Cells[assetRow, columns["AssetShare"]] = assetValue / totalValue;
            }
        }
    }

    public class Asset {
        public Asset(string name, string assetClass) {
            Name = name;
            AssetClass = assetClass;
        }
        public string Name { get; set; }
        public string AssetClass { get; set; }
    }

    public class CostBasisParameter {
        public CostBasisParameter(decimal insertionPrice, decimal investedCapital) {
            InsertionPrice = insertionPrice;
            InvestedCapital = investedCapital;
        }
        public decimal InsertionPrice { get; set; }
        public decimal InvestedCapital { get; set; }
    }
    public class AssetInvestment : Asset{
        public AssetInvestment(string name, string assetClass, decimal initialAmount, decimal currentAmount, 
            decimal insertionPrice, decimal previousCostBasis, decimal investedCapital, decimal previousSharePrice) : base(name, assetClass) {
            Name = name;
            AssetClass = assetClass;
            InitialAmount = initialAmount;
            CurrentAmount = currentAmount;
            PreviousSharePrice = previousSharePrice;
            InsertionPrice = insertionPrice != 0 ? insertionPrice : 0;
            if (currentAmount - initialAmount >= 0) {
                InvestedCapital = investedCapital != 0 ? investedCapital : previousCostBasis * initialAmount + insertionPrice * (currentAmount - initialAmount);
            } else {
                InvestedCapital = investedCapital != 0 ? investedCapital : previousCostBasis * initialAmount + previousCostBasis * (currentAmount - initialAmount);
            }
            PreviousCostBasis = previousCostBasis != 0 ? previousCostBasis : 1;
            
        }
        public decimal InitialAmount { get; set; }
        public decimal CurrentAmount { get; set; }
        public decimal PreviousSharePrice { get; set; }
        public decimal TotalValue => InsertionPrice != 0 ? InsertionPrice * CurrentAmount : PreviousSharePrice * CurrentAmount;
        public decimal InsertionPrice { get; set; }
        public decimal PreviousCostBasis { get; set; }
        public decimal CostBasis => CurrentAmount != 0 ? InvestedCapital / CurrentAmount : 0;
        public decimal InvestedCapital { get; set; }
        public decimal Performance => CurrentAmount != 0 ? (TotalValue - InvestedCapital) / InvestedCapital : 0;

        private void InsertInitialAmount(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel) {
            excel.Cells[assetRow, assetColumn] = assetInvestment.InitialAmount;
        }

        private void InsertCurrentAmount(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel) {
            excel.Cells[assetRow, assetColumn] = assetInvestment.CurrentAmount;
        }

        private void InsertTotalValue(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel) {
            excel.Cells[assetRow, assetColumn] = assetInvestment.TotalValue;
        }

        private void InsertSharePrice(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel) {
            excel.Cells[assetRow, assetColumn] = assetInvestment.InsertionPrice != 0 ? assetInvestment.InsertionPrice : assetInvestment.PreviousSharePrice;
        }

        private void InsertCostBasis(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel) {
            excel.Cells[assetRow, assetColumn] = assetInvestment.CostBasis;
        }

        private void InsertInvestedCapital(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel) {
            excel.Cells[assetRow, assetColumn] = assetInvestment.InvestedCapital;
        }

        private void InsertPerformance(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel) {
            excel.Cells[assetRow, assetColumn] = assetInvestment.Performance;
        }
        private void InsertRelativeContribution(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel) {
            var initialAmount = assetInvestment.InitialAmount;
            if (initialAmount == 0)
                excel.Cells[assetRow, assetColumn] = 1;
            else
                excel.Cells[assetRow, assetColumn] = assetInvestment.CurrentAmount / assetInvestment.InitialAmount - 1;

        }

        public void InsertAssetInvestment(AssetInvestment assetInvestment, int assetRow, Worksheet excel, Dictionary<string, string> columns, Dictionary<string, int> assets) {
            InsertInitialAmount(assetInvestment, columns["InitialAmount"], assetRow, excel);
            InsertCurrentAmount(assetInvestment, columns["Amount"], assetRow, excel);
            InsertRelativeContribution(assetInvestment, columns["RelativeContribution"], assetRow, excel);
            InsertTotalValue(assetInvestment, columns["TotalValue"], assetRow, excel);
            InsertCostBasis(assetInvestment, columns["CostBasis"], assetRow, excel);
            InsertSharePrice(assetInvestment, columns["PricePerShare"], assetRow, excel);
            InsertInvestedCapital(assetInvestment, columns["InvestedCapital"], assetRow, excel);
            InsertPerformance(assetInvestment, columns["Performance"], assetRow, excel);
            InsertAssetShare(excel, columns, assets);
        }

        public void InsertAssetShare(Worksheet excel, Dictionary<string, string> columns, Dictionary<string, int> assets) {
            const int lastRow = 150;
            foreach (var assetRow in assets.Values) {
                var assetValueCell = (Range)excel.Cells[assetRow, columns["TotalValue"]];
                decimal assetValue;
                if (assetValueCell.Value != null)
                    assetValue = (decimal) assetValueCell.Value;
                else
                    assetValue = 0;
                
                var totalValueCell = (Range)excel.Cells[lastRow, columns["TotalValue"]];
                var totalValue = (decimal)totalValueCell.Value;

                excel.Cells[assetRow, columns["AssetShare"]] = assetValue / totalValue;
            }
        }
    }
    public class AssetClass
    {
        public AssetClass(string name, int firstRow, int lastRow)
        {
            Name = name;
            FirstRow = firstRow;
            LastRow = lastRow;
        }

        public string Name { get; set; }
        public int  FirstRow { get; set; }
        public int LastRow { get; set; }

        public static List<AssetClass> GetAllAssetClasses()
        {
            var jsonData = File.ReadAllText("AssetClasses.json");
            var assetClasses = JsonConvert.DeserializeObject<List<AssetClass>>(jsonData);
            return assetClasses;
        }
        public static void UpdateAssetClassFile(List<AssetClass> assetClasses)
        {
            var content = JsonConvert.SerializeObject(assetClasses, Formatting.Indented);
            File.WriteAllText("AssetClasses.json", content);
        }

        public static Dictionary<string, int> GetAssetList()
        {
            var jsonData = File.ReadAllText("AssetList.json");
            var assetList = JsonConvert.DeserializeObject<Dictionary<string, int>>(jsonData);

            return assetList;
        }
        public static void UpdateAssetFile(Dictionary<string, int> assets)
        {
            var content = JsonConvert.SerializeObject(assets, Formatting.Indented);
            File.WriteAllText("AssetList.json", content);
        }

        public static bool AssetClassContentChanged(List<AssetClass> assetClasses, List<AssetClass> assetClassesFromFile)
        {
            var assetClassesContent = JsonConvert.SerializeObject(assetClasses, Formatting.None);
            var assetClassesContentFromFile = JsonConvert.SerializeObject(assetClassesFromFile, Formatting.None);

            var contentChanged = !string.Equals(assetClassesContent, assetClassesContentFromFile);
            return contentChanged;
        }

        public static bool AssetContentChanged(Dictionary<string, int> assets, Dictionary<string, int> assetsFromFile)
        {
            var assetsContent = JsonConvert.SerializeObject(assets, Formatting.None);
            var assetsContentFromFile = JsonConvert.SerializeObject(assetsFromFile, Formatting.None);
            
            var contentChanged = !string.Equals(assetsContent, assetsContentFromFile);
            return contentChanged;
        }

        public static bool AssetContentChanged(Dictionary<string, int> assets, List<AssetClass> assetClasses)
        {
            var assetClassesFromFile = AssetClass.GetAllAssetClasses();
            var assetClassesContentChanged = AssetClass.AssetClassContentChanged(assetClasses, assetClassesFromFile);
            var assetsFromFile = AssetClass.GetAssetList();
            var assetsContentChanged = AssetClass.AssetContentChanged(assets, assetsFromFile);
            
            if(assetClassesContentChanged || assetsContentChanged)
            {
                return true;
            }
            return false;
        }
    }
}
