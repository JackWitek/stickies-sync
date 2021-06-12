namespace NotesSyncApp
{
    partial class MainWindow
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
            this.commandTextBox = new System.Windows.Forms.TextBox();
            this.messagesTextBox = new System.Windows.Forms.TextBox();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.listBox3 = new System.Windows.Forms.ListBox();
            this.sendButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.eventsCheckBox = new System.Windows.Forms.CheckBox();
            this.button2 = new System.Windows.Forms.Button();
            this.loggingLevelDropdown = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // commandTextBox
            // 
            this.commandTextBox.Location = new System.Drawing.Point(43, 57);
            this.commandTextBox.Name = "commandTextBox";
            this.commandTextBox.Size = new System.Drawing.Size(719, 29);
            this.commandTextBox.TabIndex = 0;
            this.commandTextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.commandTextBox_KeyPress);
            // 
            // messagesTextBox
            // 
            this.messagesTextBox.BackColor = System.Drawing.SystemColors.Menu;
            this.messagesTextBox.Location = new System.Drawing.Point(43, 105);
            this.messagesTextBox.Multiline = true;
            this.messagesTextBox.Name = "messagesTextBox";
            this.messagesTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.messagesTextBox.Size = new System.Drawing.Size(876, 293);
            this.messagesTextBox.TabIndex = 1;
            // 
            // listBox1
            // 
            this.listBox1.BackColor = System.Drawing.SystemColors.Menu;
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 21;
            this.listBox1.Location = new System.Drawing.Point(36, 560);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(280, 214);
            this.listBox1.TabIndex = 2;
            // 
            // listBox2
            // 
            this.listBox2.BackColor = System.Drawing.SystemColors.Menu;
            this.listBox2.ForeColor = System.Drawing.SystemColors.WindowText;
            this.listBox2.FormattingEnabled = true;
            this.listBox2.ItemHeight = 21;
            this.listBox2.Location = new System.Drawing.Point(344, 560);
            this.listBox2.Name = "listBox2";
            this.listBox2.Size = new System.Drawing.Size(280, 214);
            this.listBox2.TabIndex = 3;
            // 
            // listBox3
            // 
            this.listBox3.BackColor = System.Drawing.SystemColors.Menu;
            this.listBox3.FormattingEnabled = true;
            this.listBox3.ItemHeight = 21;
            this.listBox3.Location = new System.Drawing.Point(651, 560);
            this.listBox3.Name = "listBox3";
            this.listBox3.Size = new System.Drawing.Size(280, 214);
            this.listBox3.TabIndex = 4;
            // 
            // sendButton
            // 
            this.sendButton.Location = new System.Drawing.Point(778, 54);
            this.sendButton.Name = "sendButton";
            this.sendButton.Size = new System.Drawing.Size(141, 32);
            this.sendButton.TabIndex = 5;
            this.sendButton.Text = "Send";
            this.sendButton.UseVisualStyleBackColor = true;
            this.sendButton.Click += new System.EventHandler(this.sendButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 13.824F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label1.ForeColor = System.Drawing.SystemColors.Control;
            this.label1.Location = new System.Drawing.Point(37, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(260, 32);
            this.label1.TabIndex = 7;
            this.label1.Text = "Custom API Command:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.SystemColors.Control;
            this.label2.Location = new System.Drawing.Point(36, 531);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 21);
            this.label2.TabIndex = 8;
            this.label2.Text = "Sent";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.SystemColors.Control;
            this.label3.Location = new System.Drawing.Point(340, 531);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(72, 21);
            this.label3.TabIndex = 8;
            this.label3.Text = "Received";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.SystemColors.Control;
            this.label4.Location = new System.Drawing.Point(651, 531);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(55, 21);
            this.label4.TabIndex = 8;
            this.label4.Text = "Events";
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.PaleGreen;
            this.button3.Location = new System.Drawing.Point(55, 427);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(190, 38);
            this.button3.TabIndex = 9;
            this.button3.Text = "Start auto sync";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.startSync_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.MistyRose;
            this.button1.Location = new System.Drawing.Point(55, 480);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(190, 38);
            this.button1.TabIndex = 10;
            this.button1.Text = "Stop";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.stopSync_Click);
            // 
            // eventsCheckBox
            // 
            this.eventsCheckBox.AutoSize = true;
            this.eventsCheckBox.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.eventsCheckBox.Location = new System.Drawing.Point(781, 435);
            this.eventsCheckBox.Name = "eventsCheckBox";
            this.eventsCheckBox.Size = new System.Drawing.Size(138, 25);
            this.eventsCheckBox.TabIndex = 11;
            this.eventsCheckBox.Text = "Retrieve Events";
            this.eventsCheckBox.UseVisualStyleBackColor = true;
            this.eventsCheckBox.CheckedChanged += new System.EventHandler(this.eventsCheckBox_CheckedChanged);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(780, 468);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(136, 29);
            this.button2.TabIndex = 12;
            this.button2.Text = "Clear Events";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.clearEvents_Click);
            // 
            // loggingLevelDropdown
            // 
            this.loggingLevelDropdown.FormattingEnabled = true;
            this.loggingLevelDropdown.Location = new System.Drawing.Point(582, 468);
            this.loggingLevelDropdown.Name = "loggingLevelDropdown";
            this.loggingLevelDropdown.Size = new System.Drawing.Size(158, 29);
            this.loggingLevelDropdown.TabIndex = 13;
            this.loggingLevelDropdown.SelectedIndexChanged += new System.EventHandler(this.loggingDropDown_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.SystemColors.Control;
            this.label5.Location = new System.Drawing.Point(580, 437);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(107, 21);
            this.label5.TabIndex = 8;
            this.label5.Text = "Logging Level";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(283, 432);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(190, 30);
            this.button4.TabIndex = 14;
            this.button4.Text = "Download from DB...";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.downloadX_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(283, 480);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(190, 30);
            this.button5.TabIndex = 14;
            this.button5.Text = "Upload to DB";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.downloadX_Click);
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.ClientSize = new System.Drawing.Size(956, 823);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.loggingLevelDropdown);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.eventsCheckBox);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.sendButton);
            this.Controls.Add(this.listBox3);
            this.Controls.Add(this.listBox2);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.messagesTextBox);
            this.Controls.Add(this.commandTextBox);
            this.Name = "MainWindow";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.MainWindow_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox commandTextBox;
        private System.Windows.Forms.TextBox messagesTextBox;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.ListBox listBox2;
        private System.Windows.Forms.ListBox listBox3;
        private System.Windows.Forms.Button sendButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox eventsCheckBox;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ComboBox loggingLevelDropdown;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
    }
}

