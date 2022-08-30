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
    public partial class AssetClassForm : Form
    {
        List<AssetClass> _assetClasses;
        public AssetClassForm(List<AssetClass> assetClasses)
        {
            _assetClasses = assetClasses;
            InitializeComponent();
            txt_firstRow.Text = CalculateFirstRowForAssetClass().ToString();
        }

        private int CalculateFirstRowForAssetClass()
        {
            var previousLastRow = _assetClasses.OrderByDescending(a => a.LastRow).First().LastRow;
            var firstRow = previousLastRow + 3; //Skip two rows for formatting reasons
            
            return firstRow;
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            Close();
        }


        private bool AllFieldsFilled()
        {
            if(txt_name.Text != "" && txt_firstRow.Text != "" && txt_classItemAmount.Text != "")
            {
                return true;
            }
            return false;
        }

        private void btn_confirm_Click(object sender, EventArgs e)
        {
            if (AllFieldsFilled())
            {
                CheckForValidIntValue(txt_firstRow);
                CheckForValidIntValue(txt_classItemAmount);

                int firstRow = Convert.ToInt32(txt_firstRow.Text);
                int lastRow = firstRow + Convert.ToInt32(txt_classItemAmount.Text);
                var assetClass = new AssetClass(txt_name.Text, firstRow, lastRow);
                _assetClasses.Add(assetClass);

                AssetClass.UpdateAssetClassFile(_assetClasses);
                Close();
            }
            else
            {
                MessageBox.Show("Not all fields were filled out!", "Missing Data", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CheckForValidIntValue(TextBox textBox)
        {
            try
            {
                Convert.ToInt32(textBox.Text);
            }
            catch (Exception exception)
            {
                Console.WriteLine("Please enter a valid number next time! Exception: " + exception.Message);
                throw;
            }
        }
    }
}
