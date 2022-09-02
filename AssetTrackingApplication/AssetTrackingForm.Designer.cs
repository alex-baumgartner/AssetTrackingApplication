namespace AssetTrackingApplication
{
    public partial class AssetTrackingForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.txt_currentSharePrice = new System.Windows.Forms.TextBox();
            this.lbl_capital = new System.Windows.Forms.Label();
            this.lb_assetName = new System.Windows.Forms.Label();
            this.txt_assetClass = new System.Windows.Forms.TextBox();
            this.btn_insertData = new System.Windows.Forms.Button();
            this.lb_filePath = new System.Windows.Forms.Label();
            this.txt_previousPrice = new System.Windows.Forms.TextBox();
            this.cb_assetName = new System.Windows.Forms.ComboBox();
            this.lb_asset_class = new System.Windows.Forms.Label();
            this.txt_currentValue = new System.Windows.Forms.TextBox();
            this.pn_OverallData = new System.Windows.Forms.Panel();
            this.btn_createAsset = new System.Windows.Forms.Button();
            this.btn_createAssetClass = new System.Windows.Forms.Button();
            this.lbl_totalValue = new System.Windows.Forms.Label();
            this.txt_totalValue = new System.Windows.Forms.TextBox();
            this.btn_finish = new System.Windows.Forms.Button();
            this.btn_InitializeAssetInsertion = new System.Windows.Forms.Button();
            this.btn_InitializeUpdate = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label7 = new System.Windows.Forms.Label();
            this.txt_percentalGain = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txt_previousSharePrice = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_priceGain = new System.Windows.Forms.TextBox();
            this.btn_updateData = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.txt_investedCapital = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txt_insertionPrice = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txt_previousAmount = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_currentAmount = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.pn_OverallData.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // txt_currentSharePrice
            // 
            this.txt_currentSharePrice.Enabled = false;
            this.txt_currentSharePrice.Location = new System.Drawing.Point(61, 69);
            this.txt_currentSharePrice.Margin = new System.Windows.Forms.Padding(4);
            this.txt_currentSharePrice.Name = "txt_currentSharePrice";
            this.txt_currentSharePrice.Size = new System.Drawing.Size(239, 22);
            this.txt_currentSharePrice.TabIndex = 0;
            // 
            // lbl_capital
            // 
            this.lbl_capital.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_capital.Location = new System.Drawing.Point(352, 34);
            this.lbl_capital.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbl_capital.Name = "lbl_capital";
            this.lbl_capital.Size = new System.Drawing.Size(244, 31);
            this.lbl_capital.TabIndex = 1;
            this.lbl_capital.Text = "Previous Value";
            this.lbl_capital.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lb_assetName
            // 
            this.lb_assetName.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_assetName.Location = new System.Drawing.Point(53, 36);
            this.lb_assetName.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lb_assetName.Name = "lb_assetName";
            this.lb_assetName.Size = new System.Drawing.Size(244, 31);
            this.lb_assetName.TabIndex = 7;
            this.lb_assetName.Text = "Asset Name";
            this.lb_assetName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txt_assetClass
            // 
            this.txt_assetClass.Location = new System.Drawing.Point(372, 73);
            this.txt_assetClass.Margin = new System.Windows.Forms.Padding(4);
            this.txt_assetClass.Name = "txt_assetClass";
            this.txt_assetClass.Size = new System.Drawing.Size(293, 22);
            this.txt_assetClass.TabIndex = 6;
            // 
            // btn_insertData
            // 
            this.btn_insertData.Enabled = false;
            this.btn_insertData.Font = new System.Drawing.Font("Microsoft Sans Serif", 27.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_insertData.Location = new System.Drawing.Point(5, 588);
            this.btn_insertData.Margin = new System.Windows.Forms.Padding(4);
            this.btn_insertData.Name = "btn_insertData";
            this.btn_insertData.Size = new System.Drawing.Size(609, 107);
            this.btn_insertData.TabIndex = 20;
            this.btn_insertData.Text = "Insert Asset";
            this.btn_insertData.UseVisualStyleBackColor = true;
            this.btn_insertData.Click += new System.EventHandler(this.btn_insertData_Click);
            // 
            // lb_filePath
            // 
            this.lb_filePath.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_filePath.Location = new System.Drawing.Point(15, 34);
            this.lb_filePath.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lb_filePath.Name = "lb_filePath";
            this.lb_filePath.Size = new System.Drawing.Size(328, 31);
            this.lb_filePath.TabIndex = 23;
            this.lb_filePath.Text = "Current Price Per Share";
            this.lb_filePath.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txt_previousPrice
            // 
            this.txt_previousPrice.Enabled = false;
            this.txt_previousPrice.Location = new System.Drawing.Point(356, 69);
            this.txt_previousPrice.Margin = new System.Windows.Forms.Padding(4);
            this.txt_previousPrice.Name = "txt_previousPrice";
            this.txt_previousPrice.Size = new System.Drawing.Size(239, 22);
            this.txt_previousPrice.TabIndex = 22;
            // 
            // cb_assetName
            // 
            this.cb_assetName.FormattingEnabled = true;
            this.cb_assetName.Location = new System.Drawing.Point(4, 71);
            this.cb_assetName.Margin = new System.Windows.Forms.Padding(4);
            this.cb_assetName.Name = "cb_assetName";
            this.cb_assetName.Size = new System.Drawing.Size(293, 24);
            this.cb_assetName.TabIndex = 24;
            this.cb_assetName.TextChanged += new System.EventHandler(this.cb_assetName_TextChanged);
            this.cb_assetName.MouseClick += new System.Windows.Forms.MouseEventHandler(this.cb_assetName_MouseClick);
            // 
            // lb_asset_class
            // 
            this.lb_asset_class.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_asset_class.Location = new System.Drawing.Point(392, 36);
            this.lb_asset_class.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lb_asset_class.Name = "lb_asset_class";
            this.lb_asset_class.Size = new System.Drawing.Size(244, 31);
            this.lb_asset_class.TabIndex = 25;
            this.lb_asset_class.Text = "Asset Class";
            this.lb_asset_class.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txt_currentValue
            // 
            this.txt_currentValue.Enabled = false;
            this.txt_currentValue.Location = new System.Drawing.Point(356, 101);
            this.txt_currentValue.Margin = new System.Windows.Forms.Padding(4);
            this.txt_currentValue.Name = "txt_currentValue";
            this.txt_currentValue.Size = new System.Drawing.Size(239, 22);
            this.txt_currentValue.TabIndex = 26;
            // 
            // pn_OverallData
            // 
            this.pn_OverallData.Controls.Add(this.btn_createAsset);
            this.pn_OverallData.Controls.Add(this.btn_createAssetClass);
            this.pn_OverallData.Controls.Add(this.lbl_totalValue);
            this.pn_OverallData.Controls.Add(this.txt_totalValue);
            this.pn_OverallData.Controls.Add(this.btn_finish);
            this.pn_OverallData.Controls.Add(this.lb_assetName);
            this.pn_OverallData.Controls.Add(this.lb_asset_class);
            this.pn_OverallData.Controls.Add(this.txt_assetClass);
            this.pn_OverallData.Controls.Add(this.cb_assetName);
            this.pn_OverallData.Location = new System.Drawing.Point(641, 15);
            this.pn_OverallData.Margin = new System.Windows.Forms.Padding(4);
            this.pn_OverallData.Name = "pn_OverallData";
            this.pn_OverallData.Size = new System.Drawing.Size(669, 702);
            this.pn_OverallData.TabIndex = 28;
            // 
            // btn_createAsset
            // 
            this.btn_createAsset.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btn_createAsset.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_createAsset.Location = new System.Drawing.Point(4, 133);
            this.btn_createAsset.Margin = new System.Windows.Forms.Padding(4);
            this.btn_createAsset.Name = "btn_createAsset";
            this.btn_createAsset.Size = new System.Drawing.Size(293, 61);
            this.btn_createAsset.TabIndex = 38;
            this.btn_createAsset.Text = "Create new Asset";
            this.btn_createAsset.UseVisualStyleBackColor = true;
            this.btn_createAsset.Click += new System.EventHandler(this.btn_createAsset_Click);
            // 
            // btn_createAssetClass
            // 
            this.btn_createAssetClass.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btn_createAssetClass.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_createAssetClass.Location = new System.Drawing.Point(372, 133);
            this.btn_createAssetClass.Margin = new System.Windows.Forms.Padding(4);
            this.btn_createAssetClass.Name = "btn_createAssetClass";
            this.btn_createAssetClass.Size = new System.Drawing.Size(293, 61);
            this.btn_createAssetClass.TabIndex = 37;
            this.btn_createAssetClass.Text = "Create new Asset Class";
            this.btn_createAssetClass.UseVisualStyleBackColor = true;
            this.btn_createAssetClass.Click += new System.EventHandler(this.btn_createAssetClass_Click);
            // 
            // lbl_totalValue
            // 
            this.lbl_totalValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_totalValue.Location = new System.Drawing.Point(227, 480);
            this.lbl_totalValue.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbl_totalValue.Name = "lbl_totalValue";
            this.lbl_totalValue.Size = new System.Drawing.Size(244, 31);
            this.lbl_totalValue.TabIndex = 36;
            this.lbl_totalValue.Text = "Total Value";
            this.lbl_totalValue.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lbl_totalValue.Visible = false;
            // 
            // txt_totalValue
            // 
            this.txt_totalValue.Enabled = false;
            this.txt_totalValue.Location = new System.Drawing.Point(232, 528);
            this.txt_totalValue.Margin = new System.Windows.Forms.Padding(4);
            this.txt_totalValue.Name = "txt_totalValue";
            this.txt_totalValue.Size = new System.Drawing.Size(239, 22);
            this.txt_totalValue.TabIndex = 35;
            // 
            // btn_finish
            // 
            this.btn_finish.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_finish.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_finish.Location = new System.Drawing.Point(4, 590);
            this.btn_finish.Margin = new System.Windows.Forms.Padding(4);
            this.btn_finish.Name = "btn_finish";
            this.btn_finish.Size = new System.Drawing.Size(661, 107);
            this.btn_finish.TabIndex = 28;
            this.btn_finish.Text = "Finish";
            this.btn_finish.UseVisualStyleBackColor = true;
            this.btn_finish.Click += new System.EventHandler(this.btn_finish_Click);
            // 
            // btn_InitializeAssetInsertion
            // 
            this.btn_InitializeAssetInsertion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_InitializeAssetInsertion.Font = new System.Drawing.Font("Microsoft Sans Serif", 27F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_InitializeAssetInsertion.Location = new System.Drawing.Point(5, 474);
            this.btn_InitializeAssetInsertion.Margin = new System.Windows.Forms.Padding(4);
            this.btn_InitializeAssetInsertion.Name = "btn_InitializeAssetInsertion";
            this.btn_InitializeAssetInsertion.Size = new System.Drawing.Size(609, 107);
            this.btn_InitializeAssetInsertion.TabIndex = 27;
            this.btn_InitializeAssetInsertion.Text = "Start Asset Insertion";
            this.btn_InitializeAssetInsertion.UseVisualStyleBackColor = true;
            this.btn_InitializeAssetInsertion.Click += new System.EventHandler(this.btn_InitializeAssetInsertion_Click);
            // 
            // btn_InitializeUpdate
            // 
            this.btn_InitializeUpdate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btn_InitializeUpdate.Font = new System.Drawing.Font("Microsoft Sans Serif", 27F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_InitializeUpdate.Location = new System.Drawing.Point(11, 474);
            this.btn_InitializeUpdate.Margin = new System.Windows.Forms.Padding(4);
            this.btn_InitializeUpdate.Name = "btn_InitializeUpdate";
            this.btn_InitializeUpdate.Size = new System.Drawing.Size(609, 107);
            this.btn_InitializeUpdate.TabIndex = 26;
            this.btn_InitializeUpdate.Text = "Start Asset Update";
            this.btn_InitializeUpdate.UseVisualStyleBackColor = true;
            this.btn_InitializeUpdate.Click += new System.EventHandler(this.btn_InitializeUpdate_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.txt_percentalGain);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.txt_previousSharePrice);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.txt_priceGain);
            this.panel1.Controls.Add(this.btn_updateData);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.btn_InitializeUpdate);
            this.panel1.Controls.Add(this.lb_filePath);
            this.panel1.Controls.Add(this.txt_currentSharePrice);
            this.panel1.Controls.Add(this.lbl_capital);
            this.panel1.Controls.Add(this.txt_previousPrice);
            this.panel1.Controls.Add(this.txt_currentValue);
            this.panel1.Location = new System.Drawing.Point(1, 15);
            this.panel1.Margin = new System.Windows.Forms.Padding(4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(631, 701);
            this.panel1.TabIndex = 29;
            // 
            // label7
            // 
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(352, 214);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(244, 31);
            this.label7.TabIndex = 33;
            this.label7.Text = "Percental Gain";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txt_percentalGain
            // 
            this.txt_percentalGain.Enabled = false;
            this.txt_percentalGain.Location = new System.Drawing.Point(356, 186);
            this.txt_percentalGain.Margin = new System.Windows.Forms.Padding(4);
            this.txt_percentalGain.Name = "txt_percentalGain";
            this.txt_percentalGain.Size = new System.Drawing.Size(239, 22);
            this.txt_percentalGain.TabIndex = 34;
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(13, 132);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(339, 31);
            this.label6.TabIndex = 32;
            this.label6.Text = "Previous Price Per Share";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txt_previousSharePrice
            // 
            this.txt_previousSharePrice.Enabled = false;
            this.txt_previousSharePrice.Location = new System.Drawing.Point(61, 101);
            this.txt_previousSharePrice.Margin = new System.Windows.Forms.Padding(4);
            this.txt_previousSharePrice.Name = "txt_previousSharePrice";
            this.txt_previousSharePrice.Size = new System.Drawing.Size(239, 22);
            this.txt_previousSharePrice.TabIndex = 31;
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(57, 214);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(244, 31);
            this.label4.TabIndex = 29;
            this.label4.Text = "Price Gain";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txt_priceGain
            // 
            this.txt_priceGain.Enabled = false;
            this.txt_priceGain.Location = new System.Drawing.Point(61, 186);
            this.txt_priceGain.Margin = new System.Windows.Forms.Padding(4);
            this.txt_priceGain.Name = "txt_priceGain";
            this.txt_priceGain.Size = new System.Drawing.Size(239, 22);
            this.txt_priceGain.TabIndex = 30;
            // 
            // btn_updateData
            // 
            this.btn_updateData.Enabled = false;
            this.btn_updateData.Font = new System.Drawing.Font("Microsoft Sans Serif", 27.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_updateData.Location = new System.Drawing.Point(11, 588);
            this.btn_updateData.Margin = new System.Windows.Forms.Padding(4);
            this.btn_updateData.Name = "btn_updateData";
            this.btn_updateData.Size = new System.Drawing.Size(609, 107);
            this.btn_updateData.TabIndex = 28;
            this.btn_updateData.Text = "Update Asset";
            this.btn_updateData.UseVisualStyleBackColor = true;
            this.btn_updateData.Click += new System.EventHandler(this.btn_updateData_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(352, 132);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(244, 31);
            this.label1.TabIndex = 27;
            this.label1.Text = "Current Value";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.txt_investedCapital);
            this.panel2.Controls.Add(this.label9);
            this.panel2.Controls.Add(this.txt_insertionPrice);
            this.panel2.Controls.Add(this.label8);
            this.panel2.Controls.Add(this.txt_previousAmount);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.txt_currentAmount);
            this.panel2.Controls.Add(this.btn_InitializeAssetInsertion);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Controls.Add(this.btn_insertData);
            this.panel2.Font = new System.Drawing.Font("Microsoft Sans Serif", 27.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.panel2.Location = new System.Drawing.Point(1319, 15);
            this.panel2.Margin = new System.Windows.Forms.Padding(4);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(621, 701);
            this.panel2.TabIndex = 30;
            // 
            // txt_investedCapital
            // 
            this.txt_investedCapital.Enabled = false;
            this.txt_investedCapital.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_investedCapital.Location = new System.Drawing.Point(325, 69);
            this.txt_investedCapital.Margin = new System.Windows.Forms.Padding(4);
            this.txt_investedCapital.Name = "txt_investedCapital";
            this.txt_investedCapital.Size = new System.Drawing.Size(239, 23);
            this.txt_investedCapital.TabIndex = 34;
            // 
            // label9
            // 
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(287, 34);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(328, 31);
            this.label9.TabIndex = 35;
            this.label9.Text = "Invested Capital";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txt_insertionPrice
            // 
            this.txt_insertionPrice.Enabled = false;
            this.txt_insertionPrice.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_insertionPrice.Location = new System.Drawing.Point(325, 101);
            this.txt_insertionPrice.Margin = new System.Windows.Forms.Padding(4);
            this.txt_insertionPrice.Name = "txt_insertionPrice";
            this.txt_insertionPrice.Size = new System.Drawing.Size(239, 23);
            this.txt_insertionPrice.TabIndex = 32;
            // 
            // label8
            // 
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(279, 132);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(328, 31);
            this.label8.TabIndex = 33;
            this.label8.Text = "Insertion Price";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txt_previousAmount
            // 
            this.txt_previousAmount.Enabled = false;
            this.txt_previousAmount.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_previousAmount.Location = new System.Drawing.Point(53, 101);
            this.txt_previousAmount.Margin = new System.Windows.Forms.Padding(4);
            this.txt_previousAmount.Name = "txt_previousAmount";
            this.txt_previousAmount.Size = new System.Drawing.Size(239, 23);
            this.txt_previousAmount.TabIndex = 28;
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(7, 132);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(328, 31);
            this.label3.TabIndex = 29;
            this.label3.Text = "Previous Amount";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txt_currentAmount
            // 
            this.txt_currentAmount.Enabled = false;
            this.txt_currentAmount.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_currentAmount.Location = new System.Drawing.Point(53, 69);
            this.txt_currentAmount.Margin = new System.Windows.Forms.Padding(4);
            this.txt_currentAmount.Name = "txt_currentAmount";
            this.txt_currentAmount.Size = new System.Drawing.Size(239, 23);
            this.txt_currentAmount.TabIndex = 24;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(7, 34);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(328, 31);
            this.label2.TabIndex = 25;
            this.label2.Text = "Current Amount";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // AssetTrackingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1956, 731);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.pn_OverallData);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "AssetTrackingForm";
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.AssetTrackingForm_FormClosing);
            this.pn_OverallData.ResumeLayout(false);
            this.pn_OverallData.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox txt_currentSharePrice;
        private System.Windows.Forms.Label lbl_capital;
        private System.Windows.Forms.Label lb_assetName;
        private System.Windows.Forms.TextBox txt_assetClass;
        private System.Windows.Forms.Button btn_insertData;
        private System.Windows.Forms.Label lb_filePath;
        private System.Windows.Forms.TextBox txt_previousPrice;
        private System.Windows.Forms.ComboBox cb_assetName;
        private System.Windows.Forms.Label lb_asset_class;
        private System.Windows.Forms.TextBox txt_currentValue;
        private System.Windows.Forms.Panel pn_OverallData;
        private System.Windows.Forms.Button btn_InitializeAssetInsertion;
        private System.Windows.Forms.Button btn_InitializeUpdate;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_currentAmount;
        private System.Windows.Forms.Button btn_updateData;
        private System.Windows.Forms.Button btn_finish;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txt_priceGain;
        private System.Windows.Forms.TextBox txt_previousAmount;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txt_previousSharePrice;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txt_percentalGain;
        private System.Windows.Forms.TextBox txt_insertionPrice;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txt_investedCapital;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label lbl_totalValue;
        private System.Windows.Forms.TextBox txt_totalValue;
        private System.Windows.Forms.Button btn_createAssetClass;
        private System.Windows.Forms.Button btn_createAsset;
    }
}

