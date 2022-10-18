using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AssetTrackingApplication
{
    public partial class AssetForm : Form
    {
        List<AssetClass> _assetClasses;
        Dictionary<string, int> _assets;
        public AssetForm(List<AssetClass> assetClasses, Dictionary<string, int> assets)
        {
            _assetClasses = assetClasses;
            _assets = assets;
            InitializeComponent();
            var assetClassNames = _assetClasses.Select(a => a.Name).ToList();
            cb_assetClasses.DataSource = assetClassNames;
        }
        public Dictionary<string, int> Assets { get; set; }


        private void btn_cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btn_confirm_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txt_name.Text))
            {
                MessageBox.Show("Not all fields were filled out - please fill out name!", "Missing Data", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var assetRow = CalculateRowForAsset();

            if (assetRow != 0)
            {
                _assets.Add(txt_name.Text, assetRow);
                _assets = _assets.OrderBy(x => x.Value).ToDictionary(x => x.Key, x => x.Value);
                Assets = _assets;
                Close();
            }
            else
            {
                MessageBox.Show("Could not insert asset - please check asset class for available rows!", "Insert asset failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private int CalculateRowForAsset()
        {
            var selectedAssetClass = _assetClasses.SingleOrDefault(a => a.Name == cb_assetClasses.Text);

            if (selectedAssetClass == null)
            {
                return 0;
            }

            var assetRowsOfClass = (from asset in _assets
                                    where asset.Value >= selectedAssetClass.FirstRow && asset.Value < selectedAssetClass.LastRow
                                    select asset.Value).ToList();

            if (assetRowsOfClass.Any())
            {
                var maxAssetRowOfClass = assetRowsOfClass.Max();
                return maxAssetRowOfClass + 1;
            }
            return selectedAssetClass.FirstRow;
        }

    }
}
