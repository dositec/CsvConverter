namespace CsvConverter
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            progressBar1 = new ProgressBar();
            groupBoxLog = new GroupBox();
            richTextBox_Log = new RichTextBox();
            groupBoxActions = new GroupBox();
            button_Personalize = new Button();
            button_Cancel = new Button();
            button_Convert = new Button();
            groupBoxArchive = new GroupBox();
            listView1 = new ListView();
            toolStrip1 = new ToolStrip();
            splitContainer1 = new SplitContainer();
            groupBoxLog.SuspendLayout();
            groupBoxActions.SuspendLayout();
            groupBoxArchive.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
            splitContainer1.Panel1.SuspendLayout();
            splitContainer1.Panel2.SuspendLayout();
            splitContainer1.SuspendLayout();
            SuspendLayout();
            // 
            // progressBar1
            // 
            progressBar1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            progressBar1.Location = new Point(168, 23);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(694, 23);
            progressBar1.TabIndex = 12;
            // 
            // groupBoxLog
            // 
            groupBoxLog.Controls.Add(richTextBox_Log);
            groupBoxLog.Dock = DockStyle.Fill;
            groupBoxLog.Location = new Point(0, 0);
            groupBoxLog.Name = "groupBoxLog";
            groupBoxLog.Size = new Size(868, 178);
            groupBoxLog.TabIndex = 16;
            groupBoxLog.TabStop = false;
            groupBoxLog.Text = "Log";
            // 
            // richTextBox_Log
            // 
            richTextBox_Log.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            richTextBox_Log.Location = new Point(6, 22);
            richTextBox_Log.Name = "richTextBox_Log";
            richTextBox_Log.ReadOnly = true;
            richTextBox_Log.Size = new Size(856, 150);
            richTextBox_Log.TabIndex = 0;
            richTextBox_Log.Text = "";
            // 
            // groupBoxActions
            // 
            groupBoxActions.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            groupBoxActions.Controls.Add(button_Personalize);
            groupBoxActions.Controls.Add(button_Cancel);
            groupBoxActions.Controls.Add(progressBar1);
            groupBoxActions.Controls.Add(button_Convert);
            groupBoxActions.Location = new Point(12, 28);
            groupBoxActions.Name = "groupBoxActions";
            groupBoxActions.Size = new Size(868, 54);
            groupBoxActions.TabIndex = 14;
            groupBoxActions.TabStop = false;
            groupBoxActions.Text = "Actions";
            // 
            // button_Personalize
            // 
            button_Personalize.Location = new Point(168, 23);
            button_Personalize.Name = "button_Personalize";
            button_Personalize.Size = new Size(90, 23);
            button_Personalize.TabIndex = 11;
            button_Personalize.Text = "Personalize";
            button_Personalize.UseVisualStyleBackColor = true;
            button_Personalize.Click += button_Personalize_Click;
            // 
            // button_Cancel
            // 
            button_Cancel.Location = new Point(87, 23);
            button_Cancel.Name = "button_Cancel";
            button_Cancel.Size = new Size(75, 23);
            button_Cancel.TabIndex = 10;
            button_Cancel.Text = "Cancel";
            button_Cancel.UseVisualStyleBackColor = true;
            button_Cancel.Click += button_Cancel_Click;
            // 
            // button_Convert
            // 
            button_Convert.Location = new Point(6, 23);
            button_Convert.Name = "button_Convert";
            button_Convert.Size = new Size(75, 23);
            button_Convert.TabIndex = 9;
            button_Convert.Text = "Convert";
            button_Convert.UseVisualStyleBackColor = true;
            button_Convert.Click += button_Convert_Click;
            // 
            // groupBoxArchive
            // 
            groupBoxArchive.Controls.Add(listView1);
            groupBoxArchive.Dock = DockStyle.Fill;
            groupBoxArchive.Location = new Point(0, 0);
            groupBoxArchive.Name = "groupBoxArchive";
            groupBoxArchive.Size = new Size(868, 322);
            groupBoxArchive.TabIndex = 13;
            groupBoxArchive.TabStop = false;
            groupBoxArchive.Text = "Files";
            // 
            // listView1
            // 
            listView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            listView1.Location = new Point(6, 22);
            listView1.Name = "listView1";
            listView1.Size = new Size(856, 294);
            listView1.TabIndex = 0;
            listView1.UseCompatibleStateImageBehavior = false;
            // 
            // toolStrip1
            // 
            toolStrip1.Location = new Point(0, 0);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new Size(892, 25);
            toolStrip1.TabIndex = 18;
            toolStrip1.Text = "toolStrip1";
            // 
            // splitContainer1
            // 
            splitContainer1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            splitContainer1.FixedPanel = FixedPanel.Panel2;
            splitContainer1.Location = new Point(12, 88);
            splitContainer1.Name = "splitContainer1";
            splitContainer1.Orientation = Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            splitContainer1.Panel1.Controls.Add(groupBoxArchive);
            // 
            // splitContainer1.Panel2
            // 
            splitContainer1.Panel2.Controls.Add(groupBoxLog);
            splitContainer1.Size = new Size(868, 504);
            splitContainer1.SplitterDistance = 322;
            splitContainer1.TabIndex = 1;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(892, 604);
            Controls.Add(splitContainer1);
            Controls.Add(toolStrip1);
            Controls.Add(groupBoxActions);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            groupBoxLog.ResumeLayout(false);
            groupBoxActions.ResumeLayout(false);
            groupBoxArchive.ResumeLayout(false);
            splitContainer1.Panel1.ResumeLayout(false);
            splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainer1).EndInit();
            splitContainer1.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private ProgressBar progressBar1;
        private GroupBox groupBoxLog;
        private RichTextBox richTextBox_Log;
        private GroupBox groupBoxActions;
        private Button button_Convert;
        private Button button_Personalize;
        private GroupBox groupBoxArchive;
        private ToolStrip toolStrip1;
        private ListView listView1;
        private Button button_Cancel;
        private SplitContainer splitContainer1;
    }
}
