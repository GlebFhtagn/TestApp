namespace TestApp
{
    partial class ComponentForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.новыйКомпонентToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.компонентToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.вложенныйКомпонентToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.переименоватьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.удалитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.отчетОСводномСоставеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ComponentsTree = new System.Windows.Forms.TreeView();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.новыйКомпонентToolStripMenuItem,
            this.переименоватьToolStripMenuItem,
            this.удалитьToolStripMenuItem,
            this.toolStripSeparator1,
            this.отчетОСводномСоставеToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(213, 98);
            // 
            // новыйКомпонентToolStripMenuItem
            // 
            this.новыйКомпонентToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.компонентToolStripMenuItem,
            this.вложенныйКомпонентToolStripMenuItem});
            this.новыйКомпонентToolStripMenuItem.Name = "новыйКомпонентToolStripMenuItem";
            this.новыйКомпонентToolStripMenuItem.Size = new System.Drawing.Size(212, 22);
            this.новыйКомпонентToolStripMenuItem.Text = "Новый Компонент";
            // 
            // компонентToolStripMenuItem
            // 
            this.компонентToolStripMenuItem.Name = "компонентToolStripMenuItem";
            this.компонентToolStripMenuItem.Size = new System.Drawing.Size(230, 22);
            this.компонентToolStripMenuItem.Text = "Компонент верхнего уровня";
            this.компонентToolStripMenuItem.Click += new System.EventHandler(this.КомпонентToolStripMenuItem_Click);
            // 
            // вложенныйКомпонентToolStripMenuItem
            // 
            this.вложенныйКомпонентToolStripMenuItem.Name = "вложенныйКомпонентToolStripMenuItem";
            this.вложенныйКомпонентToolStripMenuItem.Size = new System.Drawing.Size(230, 22);
            this.вложенныйКомпонентToolStripMenuItem.Text = "Вложенный компонент";
            this.вложенныйКомпонентToolStripMenuItem.Click += new System.EventHandler(this.вложенныйКомпонентToolStripMenuItem_Click);
            // 
            // переименоватьToolStripMenuItem
            // 
            this.переименоватьToolStripMenuItem.Name = "переименоватьToolStripMenuItem";
            this.переименоватьToolStripMenuItem.Size = new System.Drawing.Size(212, 22);
            this.переименоватьToolStripMenuItem.Text = "Переименовать";
            this.переименоватьToolStripMenuItem.Click += new System.EventHandler(this.переименоватьToolStripMenuItem_Click);
            // 
            // удалитьToolStripMenuItem
            // 
            this.удалитьToolStripMenuItem.Name = "удалитьToolStripMenuItem";
            this.удалитьToolStripMenuItem.Size = new System.Drawing.Size(212, 22);
            this.удалитьToolStripMenuItem.Text = "Удалить";
            this.удалитьToolStripMenuItem.Click += new System.EventHandler(this.удалитьToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(209, 6);
            // 
            // отчетОСводномСоставеToolStripMenuItem
            // 
            this.отчетОСводномСоставеToolStripMenuItem.Name = "отчетОСводномСоставеToolStripMenuItem";
            this.отчетОСводномСоставеToolStripMenuItem.Size = new System.Drawing.Size(212, 22);
            this.отчетОСводномСоставеToolStripMenuItem.Text = "Отчет о сводном составе";
            this.отчетОСводномСоставеToolStripMenuItem.Click += new System.EventHandler(this.отчетОСводномСоставеToolStripMenuItem_Click);
            // 
            // ComponentsTree
            // 
            this.ComponentsTree.ContextMenuStrip = this.contextMenuStrip1;
            this.ComponentsTree.Location = new System.Drawing.Point(12, 12);
            this.ComponentsTree.Name = "ComponentsTree";
            this.ComponentsTree.Size = new System.Drawing.Size(368, 319);
            this.ComponentsTree.TabIndex = 1;
            this.ComponentsTree.MouseUp += new System.Windows.Forms.MouseEventHandler(this.ComponentsTree_MouseUp);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // ComponentForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(392, 343);
            this.Controls.Add(this.ComponentsTree);
            this.Name = "ComponentForm";
            this.Text = "База Компонентов";
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem новыйКомпонентToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem компонентToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem вложенныйКомпонентToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem переименоватьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem удалитьToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem отчетОСводномСоставеToolStripMenuItem;
        private System.Windows.Forms.TreeView ComponentsTree;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}

