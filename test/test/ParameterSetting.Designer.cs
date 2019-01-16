namespace iEMS_Setting
{
    partial class ParameterSetting
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
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.label1 = new System.Windows.Forms.Label();
            this.btnRemove = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.splitContainer3 = new System.Windows.Forms.SplitContainer();
            this.btnCollectRemoveAll = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.CollectView = new System.Windows.Forms.DataGridView();
            this.Remove = new System.Windows.Forms.DataGridViewButtonColumn();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.btnParameterAll = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.ParameterView = new System.Windows.Forms.DataGridView();
            this.Add = new System.Windows.Forms.DataGridViewButtonColumn();
            this.splitContainer4 = new System.Windows.Forms.SplitContainer();
            this.splitContainer5 = new System.Windows.Forms.SplitContainer();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer3)).BeginInit();
            this.splitContainer3.Panel1.SuspendLayout();
            this.splitContainer3.Panel2.SuspendLayout();
            this.splitContainer3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.CollectView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ParameterView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer4)).BeginInit();
            this.splitContainer4.Panel1.SuspendLayout();
            this.splitContainer4.Panel2.SuspendLayout();
            this.splitContainer4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer5)).BeginInit();
            this.splitContainer5.Panel1.SuspendLayout();
            this.splitContainer5.Panel2.SuspendLayout();
            this.splitContainer5.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.splitContainer1.IsSplitterFixed = true;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.label1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.splitContainer4);
            this.splitContainer1.Panel2.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.splitContainer1.Size = new System.Drawing.Size(830, 492);
            this.splitContainer1.SplitterDistance = 25;
            this.splitContainer1.TabIndex = 0;
            this.splitContainer1.TabStop = false;
            // 
            // label1
            // 
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(830, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "label1";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnRemove
            // 
            this.btnRemove.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnRemove.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnRemove.Image = global::iEMS_Setting.Properties.Resources.Button_Left_Color_;
            this.btnRemove.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRemove.Location = new System.Drawing.Point(8, 210);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(84, 23);
            this.btnRemove.TabIndex = 3;
            this.btnRemove.Text = "Remove(R)";
            this.btnRemove.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnAdd.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnAdd.Image = global::iEMS_Setting.Properties.Resources.Button_Right_Color_;
            this.btnAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnAdd.Location = new System.Drawing.Point(8, 152);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(84, 23);
            this.btnAdd.TabIndex = 2;
            this.btnAdd.Text = "Add(A)";
            this.btnAdd.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage;
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // splitContainer3
            // 
            this.splitContainer3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer3.IsSplitterFixed = true;
            this.splitContainer3.Location = new System.Drawing.Point(0, 0);
            this.splitContainer3.Name = "splitContainer3";
            this.splitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer3.Panel1
            // 
            this.splitContainer3.Panel1.Controls.Add(this.btnCollectRemoveAll);
            this.splitContainer3.Panel1.Controls.Add(this.label3);
            // 
            // splitContainer3.Panel2
            // 
            this.splitContainer3.Panel2.Controls.Add(this.CollectView);
            this.splitContainer3.Size = new System.Drawing.Size(362, 463);
            this.splitContainer3.SplitterDistance = 25;
            this.splitContainer3.TabIndex = 1;
            this.splitContainer3.TabStop = false;
            // 
            // btnCollectRemoveAll
            // 
            this.btnCollectRemoveAll.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnCollectRemoveAll.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnCollectRemoveAll.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnCollectRemoveAll.Location = new System.Drawing.Point(265, 0);
            this.btnCollectRemoveAll.Name = "btnCollectRemoveAll";
            this.btnCollectRemoveAll.Size = new System.Drawing.Size(97, 25);
            this.btnCollectRemoveAll.TabIndex = 4;
            this.btnCollectRemoveAll.Text = "Remove All";
            this.btnCollectRemoveAll.UseVisualStyleBackColor = true;
            this.btnCollectRemoveAll.Click += new System.EventHandler(this.btnCollectRemoveAll_Click);
            // 
            // label3
            // 
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(0, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(362, 25);
            this.label3.TabIndex = 1;
            this.label3.Text = "Controllable parameter";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // CollectView
            // 
            this.CollectView.AllowUserToAddRows = false;
            this.CollectView.AllowUserToDeleteRows = false;
            this.CollectView.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            this.CollectView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.CollectView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Remove});
            this.CollectView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CollectView.Location = new System.Drawing.Point(0, 0);
            this.CollectView.Name = "CollectView";
            this.CollectView.ReadOnly = true;
            this.CollectView.RowTemplate.Height = 24;
            this.CollectView.Size = new System.Drawing.Size(362, 434);
            this.CollectView.TabIndex = 0;
            this.CollectView.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.CollectView_CellClick);
            this.CollectView.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.CollectView_CellContentClick);
            // 
            // Remove
            // 
            this.Remove.HeaderText = "Remove";
            this.Remove.Name = "Remove";
            this.Remove.ReadOnly = true;
            // 
            // splitContainer2
            // 
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.splitContainer2.IsSplitterFixed = true;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            this.splitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.btnParameterAll);
            this.splitContainer2.Panel1.Controls.Add(this.label2);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.ParameterView);
            this.splitContainer2.Size = new System.Drawing.Size(361, 463);
            this.splitContainer2.SplitterDistance = 25;
            this.splitContainer2.TabIndex = 0;
            this.splitContainer2.TabStop = false;
            // 
            // btnParameterAll
            // 
            this.btnParameterAll.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnParameterAll.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnParameterAll.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnParameterAll.Location = new System.Drawing.Point(286, 0);
            this.btnParameterAll.Name = "btnParameterAll";
            this.btnParameterAll.Size = new System.Drawing.Size(75, 25);
            this.btnParameterAll.TabIndex = 3;
            this.btnParameterAll.Text = "Add All";
            this.btnParameterAll.UseVisualStyleBackColor = true;
            this.btnParameterAll.Click += new System.EventHandler(this.btnParameterAll_Click);
            // 
            // label2
            // 
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Font = new System.Drawing.Font("新細明體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(0, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(361, 25);
            this.label2.TabIndex = 0;
            this.label2.Text = "Available parameter";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // ParameterView
            // 
            this.ParameterView.AllowUserToAddRows = false;
            this.ParameterView.AllowUserToDeleteRows = false;
            this.ParameterView.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            this.ParameterView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ParameterView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Add});
            this.ParameterView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ParameterView.Location = new System.Drawing.Point(0, 0);
            this.ParameterView.Name = "ParameterView";
            this.ParameterView.ReadOnly = true;
            this.ParameterView.RowTemplate.Height = 24;
            this.ParameterView.Size = new System.Drawing.Size(361, 434);
            this.ParameterView.TabIndex = 0;
            this.ParameterView.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.ParameterView_CellClick);
            this.ParameterView.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.ParameterView_CellContentClick);
            // 
            // Add
            // 
            this.Add.HeaderText = "Add";
            this.Add.Name = "Add";
            this.Add.ReadOnly = true;
            // 
            // splitContainer4
            // 
            this.splitContainer4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer4.IsSplitterFixed = true;
            this.splitContainer4.Location = new System.Drawing.Point(0, 0);
            this.splitContainer4.Name = "splitContainer4";
            // 
            // splitContainer4.Panel1
            // 
            this.splitContainer4.Panel1.Controls.Add(this.splitContainer2);
            // 
            // splitContainer4.Panel2
            // 
            this.splitContainer4.Panel2.Controls.Add(this.splitContainer5);
            this.splitContainer4.Size = new System.Drawing.Size(830, 463);
            this.splitContainer4.SplitterDistance = 361;
            this.splitContainer4.TabIndex = 4;
            // 
            // splitContainer5
            // 
            this.splitContainer5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer5.IsSplitterFixed = true;
            this.splitContainer5.Location = new System.Drawing.Point(0, 0);
            this.splitContainer5.Name = "splitContainer5";
            // 
            // splitContainer5.Panel1
            // 
            this.splitContainer5.Panel1.Controls.Add(this.btnAdd);
            this.splitContainer5.Panel1.Controls.Add(this.btnRemove);
            // 
            // splitContainer5.Panel2
            // 
            this.splitContainer5.Panel2.Controls.Add(this.splitContainer3);
            this.splitContainer5.Size = new System.Drawing.Size(465, 463);
            this.splitContainer5.SplitterDistance = 99;
            this.splitContainer5.TabIndex = 0;
            // 
            // ParameterSetting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(830, 492);
            this.Controls.Add(this.splitContainer1);
            this.MinimumSize = new System.Drawing.Size(846, 530);
            this.Name = "ParameterSetting";
            this.Text = "Parameter item setting";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ParameterSetting_FormClosed);
            this.Load += new System.EventHandler(this.ParameterSetting_Load);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.splitContainer3.Panel1.ResumeLayout(false);
            this.splitContainer3.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer3)).EndInit();
            this.splitContainer3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.CollectView)).EndInit();
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ParameterView)).EndInit();
            this.splitContainer4.Panel1.ResumeLayout(false);
            this.splitContainer4.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer4)).EndInit();
            this.splitContainer4.ResumeLayout(false);
            this.splitContainer5.Panel1.ResumeLayout(false);
            this.splitContainer5.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer5)).EndInit();
            this.splitContainer5.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.SplitContainer splitContainer3;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridView CollectView;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridView ParameterView;
        private System.Windows.Forms.DataGridViewButtonColumn Remove;
        private System.Windows.Forms.DataGridViewButtonColumn Add;
        private System.Windows.Forms.Button btnParameterAll;
        private System.Windows.Forms.Button btnCollectRemoveAll;
        private System.Windows.Forms.Button btnRemove;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.SplitContainer splitContainer4;
        private System.Windows.Forms.SplitContainer splitContainer5;
    }
}