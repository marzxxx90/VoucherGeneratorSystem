﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblTime = New System.Windows.Forms.Label()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.tCurrent = New System.Windows.Forms.Timer(Me.components)
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportVoucherToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ClearVoucherToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SetPrinterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CountOfAvailableVoucherToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutUsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.cboMins = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.rbOnlyOne = New System.Windows.Forms.RadioButton()
        Me.rbMass = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackgroundImage = CType(resources.GetObject("Panel1.BackgroundImage"), System.Drawing.Image)
        Me.Panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel1.Location = New System.Drawing.Point(0, 27)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1009, 393)
        Me.Panel1.TabIndex = 0
        '
        'lblTime
        '
        Me.lblTime.AutoSize = True
        Me.lblTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 30.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTime.Location = New System.Drawing.Point(503, 447)
        Me.lblTime.Name = "lblTime"
        Me.lblTime.Size = New System.Drawing.Size(410, 46)
        Me.lblTime.TabIndex = 2
        Me.lblTime.Text = "2/14/2018 9:50:20 AM"
        '
        'btnGenerate
        '
        Me.btnGenerate.Font = New System.Drawing.Font("Microsoft Sans Serif", 30.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGenerate.Location = New System.Drawing.Point(233, 426)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(264, 88)
        Me.btnGenerate.TabIndex = 0
        Me.btnGenerate.Text = "Generate"
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'tCurrent
        '
        Me.tCurrent.Enabled = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ToolsToolStripMenuItem, Me.AboutUsToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1009, 24)
        Me.MenuStrip1.TabIndex = 2
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OptionsToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'OptionsToolStripMenuItem
        '
        Me.OptionsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImportVoucherToolStripMenuItem, Me.ClearVoucherToolStripMenuItem, Me.SetPrinterToolStripMenuItem})
        Me.OptionsToolStripMenuItem.Name = "OptionsToolStripMenuItem"
        Me.OptionsToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.OptionsToolStripMenuItem.Text = "&Options"
        '
        'ImportVoucherToolStripMenuItem
        '
        Me.ImportVoucherToolStripMenuItem.Name = "ImportVoucherToolStripMenuItem"
        Me.ImportVoucherToolStripMenuItem.Size = New System.Drawing.Size(157, 22)
        Me.ImportVoucherToolStripMenuItem.Text = "&Import Voucher"
        '
        'ClearVoucherToolStripMenuItem
        '
        Me.ClearVoucherToolStripMenuItem.Name = "ClearVoucherToolStripMenuItem"
        Me.ClearVoucherToolStripMenuItem.Size = New System.Drawing.Size(157, 22)
        Me.ClearVoucherToolStripMenuItem.Text = "&Clear Voucher"
        '
        'SetPrinterToolStripMenuItem
        '
        Me.SetPrinterToolStripMenuItem.Name = "SetPrinterToolStripMenuItem"
        Me.SetPrinterToolStripMenuItem.Size = New System.Drawing.Size(157, 22)
        Me.SetPrinterToolStripMenuItem.Text = "&Set Printer"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.ExitToolStripMenuItem.Text = "&Exit"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CountOfAvailableVoucherToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(48, 20)
        Me.ToolsToolStripMenuItem.Text = "&Tools"
        '
        'CountOfAvailableVoucherToolStripMenuItem
        '
        Me.CountOfAvailableVoucherToolStripMenuItem.Name = "CountOfAvailableVoucherToolStripMenuItem"
        Me.CountOfAvailableVoucherToolStripMenuItem.Size = New System.Drawing.Size(219, 22)
        Me.CountOfAvailableVoucherToolStripMenuItem.Text = "&Count of Available Voucher"
        '
        'AboutUsToolStripMenuItem
        '
        Me.AboutUsToolStripMenuItem.Name = "AboutUsToolStripMenuItem"
        Me.AboutUsToolStripMenuItem.Size = New System.Drawing.Size(68, 20)
        Me.AboutUsToolStripMenuItem.Text = "About &Us"
        '
        'cboMins
        '
        Me.cboMins.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMins.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboMins.FormattingEnabled = True
        Me.cboMins.Location = New System.Drawing.Point(0, 473)
        Me.cboMins.Name = "cboMins"
        Me.cboMins.Size = New System.Drawing.Size(126, 32)
        Me.cboMins.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(132, 475)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(95, 25)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Minutes"
        '
        'rbOnlyOne
        '
        Me.rbOnlyOne.AutoSize = True
        Me.rbOnlyOne.Checked = True
        Me.rbOnlyOne.Location = New System.Drawing.Point(0, 447)
        Me.rbOnlyOne.Name = "rbOnlyOne"
        Me.rbOnlyOne.Size = New System.Drawing.Size(69, 17)
        Me.rbOnlyOne.TabIndex = 0
        Me.rbOnlyOne.TabStop = True
        Me.rbOnlyOne.Text = "Only One"
        Me.rbOnlyOne.UseVisualStyleBackColor = True
        '
        'rbMass
        '
        Me.rbMass.AutoSize = True
        Me.rbMass.Location = New System.Drawing.Point(88, 447)
        Me.rbMass.Name = "rbMass"
        Me.rbMass.Size = New System.Drawing.Size(61, 17)
        Me.rbMass.TabIndex = 7
        Me.rbMass.Text = "Multiple"
        Me.rbMass.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(-4, 426)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 20)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Generate"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.ClientSize = New System.Drawing.Size(1009, 517)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.rbMass)
        Me.Controls.Add(Me.rbOnlyOne)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboMins)
        Me.Controls.Add(Me.lblTime)
        Me.Controls.Add(Me.btnGenerate)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmMain"
        Me.Text = "Voucher Generator System"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnGenerate As System.Windows.Forms.Button
    Friend WithEvents lblTime As System.Windows.Forms.Label
    Friend WithEvents tCurrent As System.Windows.Forms.Timer
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportVoucherToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ClearVoucherToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SetPrinterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutUsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CountOfAvailableVoucherToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cboMins As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents rbOnlyOne As System.Windows.Forms.RadioButton
    Friend WithEvents rbMass As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
