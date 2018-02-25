namespace SharedAddInEvents
{
    partial class SimpleUserControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnHello = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnHello
            // 
            this.btnHello.Font = new System.Drawing.Font("¸¼Àº °íµñ", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.btnHello.Location = new System.Drawing.Point(5, 4);
            this.btnHello.Name = "btnHello";
            this.btnHello.Size = new System.Drawing.Size(95, 26);
            this.btnHello.TabIndex = 0;
            this.btnHello.Text = "Say Hello";
            this.btnHello.UseVisualStyleBackColor = true;
            this.btnHello.Click += new System.EventHandler(this.btnHello_Click);
            // 
            // SimpleUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnHello);
            this.Name = "SimpleUserControl";
            this.Size = new System.Drawing.Size(194, 196);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnHello;
    }
}
