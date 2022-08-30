
namespace AssetTrackingApplication
{
    partial class AssetClassForm
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
            this.btn_cancel = new System.Windows.Forms.Button();
            this.btn_confirm = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lbl_Range = new System.Windows.Forms.Label();
            this.lbl_lastRow = new System.Windows.Forms.Label();
            this.lbl_firstRow = new System.Windows.Forms.Label();
            this.lbl_name = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.txt_classItemAmount = new System.Windows.Forms.TextBox();
            this.txt_firstRow = new System.Windows.Forms.TextBox();
            this.txt_name = new System.Windows.Forms.TextBox();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_cancel
            // 
            this.btn_cancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_cancel.Location = new System.Drawing.Point(244, 294);
            this.btn_cancel.Name = "btn_cancel";
            this.btn_cancel.Size = new System.Drawing.Size(121, 31);
            this.btn_cancel.TabIndex = 0;
            this.btn_cancel.Text = "Cancel";
            this.btn_cancel.UseVisualStyleBackColor = true;
            this.btn_cancel.Click += new System.EventHandler(this.btn_cancel_Click);
            // 
            // btn_confirm
            // 
            this.btn_confirm.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btn_confirm.Location = new System.Drawing.Point(13, 294);
            this.btn_confirm.Name = "btn_confirm";
            this.btn_confirm.Size = new System.Drawing.Size(121, 31);
            this.btn_confirm.TabIndex = 1;
            this.btn_confirm.Text = "Confirm";
            this.btn_confirm.UseVisualStyleBackColor = true;
            this.btn_confirm.Click += new System.EventHandler(this.btn_confirm_Click);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.panel1.Controls.Add(this.lbl_Range);
            this.panel1.Controls.Add(this.lbl_lastRow);
            this.panel1.Controls.Add(this.lbl_firstRow);
            this.panel1.Controls.Add(this.lbl_name);
            this.panel1.Location = new System.Drawing.Point(13, 13);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(171, 275);
            this.panel1.TabIndex = 2;
            // 
            // lbl_Range
            // 
            this.lbl_Range.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Range.Location = new System.Drawing.Point(1, 126);
            this.lbl_Range.Name = "lbl_Range";
            this.lbl_Range.Size = new System.Drawing.Size(115, 30);
            this.lbl_Range.TabIndex = 3;
            this.lbl_Range.Text = "Range";
            // 
            // lbl_lastRow
            // 
            this.lbl_lastRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_lastRow.Location = new System.Drawing.Point(3, 216);
            this.lbl_lastRow.Name = "lbl_lastRow";
            this.lbl_lastRow.Size = new System.Drawing.Size(165, 23);
            this.lbl_lastRow.TabIndex = 2;
            this.lbl_lastRow.Text = "Amount of Items:";
            // 
            // lbl_firstRow
            // 
            this.lbl_firstRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_firstRow.Location = new System.Drawing.Point(3, 175);
            this.lbl_firstRow.Name = "lbl_firstRow";
            this.lbl_firstRow.Size = new System.Drawing.Size(115, 23);
            this.lbl_firstRow.TabIndex = 1;
            this.lbl_firstRow.Text = "First Row:";
            // 
            // lbl_name
            // 
            this.lbl_name.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_name.Location = new System.Drawing.Point(2, 31);
            this.lbl_name.Name = "lbl_name";
            this.lbl_name.Size = new System.Drawing.Size(115, 23);
            this.lbl_name.TabIndex = 0;
            this.lbl_name.Text = "Name:";
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.panel2.Controls.Add(this.txt_classItemAmount);
            this.panel2.Controls.Add(this.txt_firstRow);
            this.panel2.Controls.Add(this.txt_name);
            this.panel2.Location = new System.Drawing.Point(190, 13);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(175, 275);
            this.panel2.TabIndex = 3;
            // 
            // txt_classItemAmount
            // 
            this.txt_classItemAmount.Location = new System.Drawing.Point(7, 216);
            this.txt_classItemAmount.Name = "txt_classItemAmount";
            this.txt_classItemAmount.Size = new System.Drawing.Size(165, 22);
            this.txt_classItemAmount.TabIndex = 6;
            // 
            // txt_firstRow
            // 
            this.txt_firstRow.Location = new System.Drawing.Point(7, 175);
            this.txt_firstRow.Name = "txt_firstRow";
            this.txt_firstRow.Size = new System.Drawing.Size(165, 22);
            this.txt_firstRow.TabIndex = 5;
            // 
            // txt_name
            // 
            this.txt_name.Location = new System.Drawing.Point(7, 31);
            this.txt_name.Name = "txt_name";
            this.txt_name.Size = new System.Drawing.Size(165, 22);
            this.txt_name.TabIndex = 4;
            // 
            // AssetClassForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(377, 337);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btn_confirm);
            this.Controls.Add(this.btn_cancel);
            this.Name = "AssetClassForm";
            this.Text = "Add Asset Class";
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn_cancel;
        private System.Windows.Forms.Button btn_confirm;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lbl_name;
        private System.Windows.Forms.Label lbl_lastRow;
        private System.Windows.Forms.Label lbl_firstRow;
        private System.Windows.Forms.Label lbl_Range;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TextBox txt_name;
        private System.Windows.Forms.TextBox txt_classItemAmount;
        private System.Windows.Forms.TextBox txt_firstRow;
    }
}