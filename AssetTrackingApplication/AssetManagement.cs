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
        public AssetUpdate(string name, string assetClass, decimal pricePerShare, decimal amount, decimal previousValue, decimal investedCapital) : base(name, assetClass) 
        {
            Name = name;
            AssetClass = assetClass;
            PricePerShare = pricePerShare;
            Amount = amount;
            PreviousValue = previousValue;
            InvestedCapital = investedCapital;
        }

        public decimal PricePerShare { get; set; }
        public decimal Amount { get; set; }
        public decimal TotalValue => (PricePerShare != 0 ? PricePerShare : TotalValue) * Amount;
        public decimal PreviousValue { get; set; }
        public decimal GainTotal => TotalValue - PreviousValue;
        public decimal GainRelative => PreviousValue != 0 ? ((TotalValue / PreviousValue) - 1) : TotalValue == 0 ? 0 : 1;
        public decimal InvestedCapital { get; set; }
        public decimal Performance => InvestedCapital != 0 ? (TotalValue - InvestedCapital) / InvestedCapital :0;

        public void InsertAssetUpdate(AssetUpdate assetUpdate, int assetRow, Worksheet excel, Dictionary<string, string> columns, Dictionary<string, int> assets)
        {
            InsertAssetName(assetUpdate, columns["Name"], assetRow, excel);
            InsertAssetClass(assetUpdate, columns["AssetClass"], assetRow, excel);
            InsertPricePerShare(assetUpdate, columns["PricePerShare"], assetRow, excel);
            InsertTotalValue(assetUpdate, columns["TotalValue"], assetRow, excel);
            InsertPreviousValue(assetUpdate, columns["PreviousValue"], assetRow, excel);
            InsertGainTotal(assetUpdate, columns["GainTotal"], assetRow, excel);
            InsertGainRelative(assetUpdate, columns["GainRelative"], assetRow, excel);
            InsertPerformance(assetUpdate, columns["Performance"], assetRow, excel);
            InsertAssetShare(excel, columns, assets);
        }

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
            if(assetUpdate.AssetClass == "Bank Accounts")
            {
                return;
            }

            excel.Cells[assetRow, assetColumn] = assetUpdate.Performance;
        }

        private void InsertAssetClass(AssetUpdate assetUpdate, string assetColumn, int assetRow, Worksheet excel)
        {
            excel.Cells[assetRow, assetColumn] = assetUpdate.AssetClass;
        }

        private void InsertAssetName(AssetUpdate assetUpdate, string assetColumn, int assetRow, Worksheet excel)
        {
            excel.Cells[assetRow, assetColumn] = assetUpdate.Name;
        }

        public void InsertAssetShare(Worksheet excel, Dictionary<string, string> columns, Dictionary<string, int> assets) {
            const int lastRow = 150;
            foreach (var assetRow in assets.Values) 
            {
                var assetValueCell = (Range)excel.Cells[assetRow, columns["TotalValue"]];
                decimal assetValue;
                if (assetValueCell.Value != null && assetValueCell.Value != 0)
                {
                    assetValue = (decimal)assetValueCell.Value;
                    var totalValueCell = (Range)excel.Cells[lastRow, columns["TotalValue"]];
                    var totalValue = (decimal)totalValueCell.Value;
                    excel.Cells[assetRow, columns["AssetShare"]] = assetValue / (totalValue != 0 ? totalValue : assetValue);
                }
            }
        }
    }

    public class Asset 
    {
        public Asset(string name, string assetClass) 
        {
            Name = name;
            AssetClass = assetClass;
        }
        public string Name { get; set; }
        public string AssetClass { get; set; }
    }

    public class CostBasisParameter 
    {
        public CostBasisParameter(decimal insertionPrice, decimal investedCapital) 
        {
            InsertionPrice = insertionPrice;
            InvestedCapital = investedCapital;
        }
        public decimal InsertionPrice { get; set; }
        public decimal InvestedCapital { get; set; }
    }

    public class AssetInvestment : Asset 
    {
        public AssetInvestment(string name, string assetClass, decimal initialAmount, decimal currentAmount,
            decimal insertionPrice, decimal previousCostBasis, decimal investedCapital, decimal previousSharePrice) : base(name, assetClass) 
        {
            Name = name;
            AssetClass = assetClass;
            InitialAmount = initialAmount;
            CurrentAmount = currentAmount;
            PreviousSharePrice = previousSharePrice;
            InsertionPrice = insertionPrice != 0 ? insertionPrice : 0;
            if (currentAmount - initialAmount >= 0)
            {
                InvestedCapital = investedCapital != 0 ? investedCapital : previousCostBasis * initialAmount + insertionPrice * (currentAmount - initialAmount);
            }
            else
            {
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


        private void InsertAssetClass(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel)
        {
            excel.Cells[assetRow, assetColumn] = assetInvestment.AssetClass;
        }

        private void InsertAssetName(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel)
        {
            excel.Cells[assetRow, assetColumn] = assetInvestment.Name;
        }
        private void InsertInitialAmount(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel)
        {
            excel.Cells[assetRow, assetColumn] = assetInvestment.InitialAmount;
        }

        private void InsertCurrentAmount(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel)
        {
            excel.Cells[assetRow, assetColumn] = assetInvestment.CurrentAmount;
        }

        private void InsertTotalValue(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel)
        {
            excel.Cells[assetRow, assetColumn] = assetInvestment.TotalValue;
        }

        private void InsertSharePrice(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel)
        {
            excel.Cells[assetRow, assetColumn] = assetInvestment.InsertionPrice != 0 ? assetInvestment.InsertionPrice : assetInvestment.PreviousSharePrice;
        }

        private void InsertCostBasis(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel)
        {
            if(assetInvestment.AssetClass == "Bank Accounts")
            {
                return;
            }

            excel.Cells[assetRow, assetColumn] = assetInvestment.CostBasis;
        }

        private void InsertInvestedCapital(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel)
        {
            excel.Cells[assetRow, assetColumn] = assetInvestment.InvestedCapital;
        }

        private void InsertPerformance(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel)
        {
            if(assetInvestment.AssetClass == "Bank Accounts")
            {
                return;
            }

            excel.Cells[assetRow, assetColumn] = assetInvestment.Performance;
        }
        private void InsertRelativeContribution(AssetInvestment assetInvestment, string assetColumn, int assetRow, Worksheet excel)
        {
            var initialAmount = assetInvestment.InitialAmount;
            if (initialAmount == 0)
                excel.Cells[assetRow, assetColumn] = 1;
            else
                excel.Cells[assetRow, assetColumn] = assetInvestment.CurrentAmount / assetInvestment.InitialAmount - 1;

        }

        public void InsertAssetInvestment(AssetInvestment assetInvestment, int assetRow, Worksheet excel, Dictionary<string, string> columns, Dictionary<string, int> assets)
        {
            InsertAssetName(assetInvestment, columns["Name"], assetRow, excel);
            InsertAssetClass(assetInvestment, columns["AssetClass"], assetRow, excel);
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

        public void InsertAssetShare(Worksheet excel, Dictionary<string, string> columns, Dictionary<string, int> assets)
        {
            const int lastRow = 150;
            foreach (var assetRow in assets.Values)
            {
                var assetValueCell = (Range)excel.Cells[assetRow, columns["TotalValue"]];
                decimal assetValue;
                if (assetValueCell.Value != null && assetValueCell.Value != 0)
                {
                    assetValue = (decimal)assetValueCell.Value;
                    var totalValueCell = (Range)excel.Cells[lastRow, columns["TotalValue"]];
                    var totalValue = (decimal)totalValueCell.Value;
                    excel.Cells[assetRow, columns["AssetShare"]] = assetValue / (totalValue != 0 ? totalValue : assetValue);
                }
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

        public static List<AssetClass> GetAllAssetClasses(string path)
        {
            var jsonData = File.ReadAllText(path + "/AssetClasses.json");
            var assetClasses = JsonConvert.DeserializeObject<List<AssetClass>>(jsonData);
            return assetClasses;
        }
        public static void UpdateAssetClassFile(List<AssetClass> assetClasses, string path)
        {
            var content = JsonConvert.SerializeObject(assetClasses, Formatting.Indented);
            File.WriteAllText(path + "/AssetClasses.json", content);
        }

        public static Dictionary<string, int> GetAssetList(string path)
        {
            var jsonData = File.ReadAllText(path + "/AssetList.json");
            var assetList = JsonConvert.DeserializeObject<Dictionary<string, int>>(jsonData);

            return assetList;
        }
        public static void UpdateAssetFile(Dictionary<string, int> assets, string path)
        {
            var content = JsonConvert.SerializeObject(assets, Formatting.Indented);
            File.WriteAllText(path + "/AssetList.json", content);
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

        public static bool AssetContentChanged(Dictionary<string, int> assets, List<AssetClass> assetClasses, string path)
        {
            var assetClassesFromFile = GetAllAssetClasses(path);
            var assetClassesContentChanged = AssetClassContentChanged(assetClasses, assetClassesFromFile);
            var assetsFromFile = GetAssetList(path);
            var assetsContentChanged = AssetContentChanged(assets, assetsFromFile);
            
            if(assetClassesContentChanged || assetsContentChanged)
            {
                return true;
            }
            return false;
        }
    }
}
