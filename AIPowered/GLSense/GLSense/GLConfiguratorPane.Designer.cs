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
            //
            // GLConfiguratorPane
            //
            this.ClientSize = new System.Drawing.Size(600, 500);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "GLConfiguratorPane";
            this.Text = "GLConfiguratorPane";
            this.ADXBeforeTaskPaneShow += new AddinExpress.XL.ADXBeforeTaskPaneShowEventHandler(this.GLConfiguratorPane_ADXBeforeTaskPaneShow);
            this.Resize += new System.EventHandler(this.GLConfiguratorPane_Resize);
        }
        #endregion
    }
}
