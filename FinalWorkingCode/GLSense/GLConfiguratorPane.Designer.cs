namespace GLSense
{
    partial class GLConfiguratorPane
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
  
        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }
  
        #region Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GLConfiguratorPane));
            this.SuspendLayout();
            // 
            // GLConfiguratorPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(312, 500);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "GLConfiguratorPane";
            this.Text = " ";
            this.ADXBeforeTaskPaneShow += new AddinExpress.XL.ADXBeforeTaskPaneShowEventHandler(this.GLConfiguratorPane_ADXBeforeTaskPaneShow);
            this.Resize += new System.EventHandler(this.GLConfiguratorPane_Resize);
            this.ResumeLayout(false);

        }
        #endregion
    }
}
