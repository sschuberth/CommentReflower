// Comment Reflower Setting Dialog
// Copyright (C) 2004  Ian Nowland
// 
// This program is free software; you can redistribute it and/or
// modify it under the terms of the GNU General Public License
// as published by the Free Software Foundation; either version 2
// of the License, or (at your option) any later version.
// 
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program; if not, write to the Free Software
// Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

using System;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Collections.Specialized;
using System.Security.Permissions;
using Microsoft.Win32;
using CommentReflowerLib;


namespace CommentReflower
{
    /// <summary>
    /// Summary description for CommentReflowerSetup.
    /// </summary>
    class CommentReflowerSetup : System.Windows.Forms.Form
    {
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage GeneralTab;
        private System.Windows.Forms.TabPage BlocksTab;
        private System.Windows.Forms.TabPage BulletsTab;
        private System.Windows.Forms.TabPage BreakFlowStringsTab;
        private System.Windows.Forms.Button OkBtn;
        private System.Windows.Forms.Button CancelBtn;
        private System.Windows.Forms.Button HelpBtn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListView BlockList;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox BlockEndTypeCombo;
        private System.Windows.Forms.ComboBox BlockStartTypeCombo;
        private System.Windows.Forms.TextBox BlockEndText;
        private System.Windows.Forms.TextBox BlockStartText;
        private System.Windows.Forms.TextBox BulletStringText;
        private System.Windows.Forms.Button BulletNewButton;
        private System.Windows.Forms.Button BulletUpButton;
        private System.Windows.Forms.Button BulletDownButton;
        private System.Windows.Forms.Button BulletDeleteButton;
        private System.Windows.Forms.ListView BulletList;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox BulletEdgeCombo;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.CheckBox UseTabsToIndentCheck;
        private System.Windows.Forms.CheckBox BreakFlowStringIsRegExCheck;
        private System.Windows.Forms.TextBox BreakFlowStringStringText;
        private System.Windows.Forms.Button BreakFlowStringNewButton;
        private System.Windows.Forms.Button BreakFlowStringUpButton;
        private System.Windows.Forms.Button BreakFlowStringDownButton;
        private System.Windows.Forms.Button BreakFlowStringDeleteButton;
        private System.Windows.Forms.ListView BreakFlowStringList;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.CheckBox BulletIsRegExCheck;
        private System.Windows.Forms.Button BlockNewButton;
        private System.Windows.Forms.Button BlockUpButton;
        private System.Windows.Forms.Button BlockDownButton;
        private System.Windows.Forms.Button BlockDeleteButton;
        private System.Windows.Forms.TextBox BlockLineStartText;
        private System.Windows.Forms.ComboBox BlockFileTypeCombo;
        private System.Windows.Forms.CheckBox IsBlockLineStartRegExCheck;
        private System.Windows.Forms.CheckBox IsBlockEndRegExCheck;
        private System.Windows.Forms.CheckBox IsBlockStartRegExCheck;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.ComboBox FirstLineBlockCombo;
        private System.Windows.Forms.CheckBox BreakFlowStringFlowPreviousCheck;
        private System.Windows.Forms.CheckBox BreakFlowStringFlowNextCheck;
        private System.Windows.Forms.TextBox BlockMinimumWidthText;
        private System.Windows.Forms.TextBox BlockWrapWidthText;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button AboutBtn;

        /** the ParameterSet set by the dialog */
        public ParameterSet mpset;

        public CommentReflowerSetup(ParameterSet pset)
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            mpset = new ParameterSet(pset);

            UseTabsToIndentCheck.Checked = mpset.mUseTabsToIndent;
            BlockWrapWidthText.Text = mpset.mWrapWidth.ToString();
            BlockMinimumWidthText.Text = mpset.mMinimumBlockWidth.ToString();

            
            foreach (CommentBlock cb in mpset.mCommentBlocks)
            {
                BlockList.Items.Add(new ListViewItem(cb.mName));
                if (!mAllFileAssociations.Contains(cb.getAssociationsAsString()))
                {
                    mAllFileAssociations.Add(cb.getAssociationsAsString());
                }
            }
            if (BlockList.Items.Count > 0)
            {
                selectBlockListItem(0);
                updateItemsForBlock(0);
            }
            else
            {
                updateItemsForBlock(-1);
            }
            foreach (string st in mAllFileAssociations)
            {
                BlockFileTypeCombo.Items.Add(st);
            }

            foreach (BulletPoint bp in mpset.mBulletPoints)
            {
                BulletList.Items.Add(new ListViewItem(bp.mName));
            }
            if (BulletList.Items.Count > 0)
            {
                selectBulletListItem(0);
                updateItemsForBullet(0);
            }
            else
            {
                updateItemsForBullet(-1);
            }

            foreach (BreakFlowString lb in mpset.mBreakFlowStrings)
            {
                BreakFlowStringList.Items.Add(new ListViewItem(lb.mName));
            }
            if (BreakFlowStringList.Items.Count > 0)
            {
                selectBreakFlowStringListItem(0);
                updateItemsForBreakFlowString(0);
            }
            else
            {
                updateItemsForBreakFlowString(-1);
            }

        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose( bool disposing )
        {
            if( disposing )
            {
                if(components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose( disposing );
        }

        private bool mFirstBlockListError = true;
        private int mCurrentBlockListItemSelected = -1;
        private int mIndexToSelectOnBlockListMouseUp = -1;
        private bool mProgramaticallySettingSelectedBlock = false;

        private static StringCollection mAllFileAssociations = new StringCollection();

        private void updateItemsForBlock(int index)
        {
            BlockNewButton.Enabled = true;
            if (index == -1)
            {
                BlockDeleteButton.Enabled = false;
                BlockDownButton.Enabled = false;
                BlockUpButton.Enabled = false;

                BlockFileTypeCombo.Enabled = false;
                BlockFileTypeCombo.Text = "";
                BlockStartText.Enabled = false;
                BlockStartText.Text = "";
                BlockEndText.Enabled = false;
                BlockEndText.Text = "";
                BlockLineStartText.Enabled = false;
                BlockLineStartText.Text = "";
                IsBlockStartRegExCheck.Enabled = false;
                IsBlockStartRegExCheck.Checked = false;
                IsBlockEndRegExCheck.Enabled = false;
                IsBlockEndRegExCheck.Checked = false;
                IsBlockLineStartRegExCheck.Enabled = false;
                IsBlockLineStartRegExCheck.Checked = false;
                BlockStartTypeCombo.Enabled = false;
                BlockStartTypeCombo.Text = "";
                BlockEndTypeCombo.Enabled = false;
                BlockEndTypeCombo.Text = "";
                FirstLineBlockCombo.Enabled = false;
                FirstLineBlockCombo.Text = "";
                BlockWrapWidthText.Enabled = false;
                BlockWrapWidthText.Text = "";
                BlockMinimumWidthText.Enabled = false;
                BlockMinimumWidthText.Text = "";
            }
            else
            {
                BlockDeleteButton.Enabled = true;
                BlockUpButton.Enabled = (index > 0);
                BlockDownButton.Enabled = (index < (BlockList.Items.Count-1));

                CommentBlock cb = (CommentBlock)mpset.mCommentBlocks[index];

                BlockFileTypeCombo.Enabled = true;
                BlockFileTypeCombo.Text = cb.getAssociationsAsString();

                BlockStartTypeCombo.Enabled = true;
                BlockStartTypeCombo.SelectedIndex = (int)cb.mBlockStartType;
                if (BlockStartTypeCombo.SelectedIndex != 0)
                {
                    BlockStartText.Enabled = true;
                    BlockStartText.Text = convertToDisplay(cb.mBlockStart,cb.mIsBlockStartRegEx);
                    IsBlockStartRegExCheck.Enabled = true;
                    IsBlockStartRegExCheck.Checked = cb.mIsBlockStartRegEx;
                }
                else
                {
                    BlockStartText.Enabled = false;
                    BlockStartText.Text = "";
                    IsBlockStartRegExCheck.Enabled = false;
                    IsBlockStartRegExCheck.Checked = false;
                }
                BlockEndTypeCombo.Enabled = true;
                BlockEndTypeCombo.SelectedIndex = (int)cb.mBlockEndType;
                if (BlockEndTypeCombo.SelectedIndex != 0)
                {
                    BlockEndText.Enabled = true;
                    BlockEndText.Text = convertToDisplay(cb.mBlockEnd,cb.mIsBlockEndRegEx);
                    IsBlockEndRegExCheck.Enabled = true;
                    IsBlockEndRegExCheck.Checked = cb.mIsBlockEndRegEx;
                }
                else
                {
                    BlockEndText.Enabled = false;
                    BlockEndText.Text = "";
                    IsBlockEndRegExCheck.Enabled = false;
                    IsBlockEndRegExCheck.Checked = false;
                }
                BlockLineStartText.Enabled = true;
                BlockLineStartText.Text = convertToDisplay(cb.mLineStart,false);
                IsBlockLineStartRegExCheck.Enabled = false;
                IsBlockLineStartRegExCheck.Checked = false;
                FirstLineBlockCombo.Enabled = true;
                if (cb.mOnlyEmptyLineBeforeStartOfBlock)
                {
                    FirstLineBlockCombo.SelectedIndex = 0;
                }
                else
                {
                    FirstLineBlockCombo.SelectedIndex = 1;
                }

            }
        }

        private void selectBlockListItem(int index)
        {
            mCurrentBlockListItemSelected = index;
            if (index != -1)
            {
                mProgramaticallySettingSelectedBlock = true;
                BlockList.Items[index].Selected = true;
                mProgramaticallySettingSelectedBlock = false;
            }
        }

        private bool validateSelectedBlock(bool printError)
        {
            int index = mCurrentBlockListItemSelected;
            if (index != -1)
            {
                CommentBlock cb = (CommentBlock)mpset.mCommentBlocks[index];
                try
                {
                    cb.mName = BlockList.Items[index].Text;

                    cb.mFileAssociations = CommentBlock.createFileAssocFromString(BlockFileTypeCombo.Text);
                    if (!mAllFileAssociations.Contains(cb.getAssociationsAsString()))
                    {
                        mAllFileAssociations.Add(cb.getAssociationsAsString());
                        BlockFileTypeCombo.Items.Add(cb.getAssociationsAsString());
                    }
                    BlockFileTypeCombo.Text = cb.getAssociationsAsString();


                    cb.mIsBlockStartRegEx = IsBlockStartRegExCheck.Checked;
                    cb.mIsBlockEndRegEx = IsBlockEndRegExCheck.Checked;

                    cb.mBlockStart = convertFromDisplay(BlockStartText.Text,cb.mIsBlockStartRegEx);
                    cb.mBlockEnd = convertFromDisplay(BlockEndText.Text,cb.mIsBlockEndRegEx);
                    cb.mLineStart = convertFromDisplay(BlockLineStartText.Text,false);

                    cb.mOnlyEmptyLineBeforeStartOfBlock = FirstLineBlockCombo.SelectedIndex == 0;
                    BlockStartTypeCombo.Text = "";
                    BlockEndTypeCombo.Text = "";

                    mpset.validateCommentBlock(index);
                }
                catch (Exception e)
                {
                    if (printError)
                    {
                        MessageBox.Show("Error detected validating Block - " + e.Message);
                    }
                    return false;
                }
            }
            return true;
        }

        private bool mFirstBulletListError = true;
        private int mCurrentBulletListItemSelected = -1;
        private int mIndexToSelectOnBulletListMouseUp = -1;
        private bool mProgramaticallySettingSelectedBullet = false;

        private void updateItemsForBullet(int index)
        {
            BulletNewButton.Enabled = true;
            if (index == -1)
            {
                BulletDeleteButton.Enabled = false;
                BulletDownButton.Enabled = false;
                BulletUpButton.Enabled = false;

                BulletStringText.Text = "";
                BulletStringText.Enabled = false;
                BulletIsRegExCheck.Checked = false;
                BulletIsRegExCheck.Enabled = false;
                BulletEdgeCombo.Text = "";
                BulletEdgeCombo.Enabled = false;
            }
            else
            {
                BulletDeleteButton.Enabled = true;
                BulletUpButton.Enabled = (index > 0);
                BulletDownButton.Enabled = (index < (BulletList.Items.Count-1));

                BulletPoint bp = (BulletPoint)mpset.mBulletPoints[index];

                BulletStringText.Enabled = true;
                BulletStringText.Text = convertToDisplay(bp.mString,bp.mIsRegEx);
                BulletEdgeCombo.Enabled = true;
                if (bp.mWrapIsAtRight)
                {
                    BulletEdgeCombo.Text = "Right";
                }
                else
                {
                    BulletEdgeCombo.Text = "Left";
                }

                BulletIsRegExCheck.Enabled = true;
                BulletIsRegExCheck.Checked = bp.mIsRegEx;
            }
        }

        private void selectBulletListItem(int index)
        {
            mCurrentBulletListItemSelected = index;
            if (index != -1)
            {
                mProgramaticallySettingSelectedBullet = true;
                BulletList.Items[index].Selected = true;
                mProgramaticallySettingSelectedBullet = false;
            }
        }

        private bool validateSelectedBullet(bool printError)
        {
            int index = mCurrentBulletListItemSelected;
            if (index != -1)
            {
                BulletPoint bp = (BulletPoint)mpset.mBulletPoints[index];
                bp.mName = BulletList.Items[index].Text;
                bp.mIsRegEx = BulletIsRegExCheck.Checked;
                bp.mString = convertFromDisplay(BulletStringText.Text,bp.mIsRegEx);
                bp.mWrapIsAtRight = BulletEdgeCombo.Text == "Right";
                try
                {
                    mpset.validateBullet(index);
                }
                catch (Exception e)
                {
                    if (printError)
                    {
                        MessageBox.Show("Error detected validating bullet - " + e.Message);
                    }
                    return false;
                }
            }
            return true;
        }

        private bool mFirstBreakFlowStringListError = true;
        private int mCurrentBreakFlowStringListItemSelected = -1;
        private int mIndexToSelectOnBreakFlowStringListMouseUp = -1;
        private bool mProgramaticallySettingSelectedBreakFlowString = false;

        private void updateItemsForBreakFlowString(int index)
        {
            BreakFlowStringNewButton.Enabled = true;
            if (index == -1)
            {
                BreakFlowStringDeleteButton.Enabled = false;
                BreakFlowStringDownButton.Enabled = false;
                BreakFlowStringUpButton.Enabled = false;

                BreakFlowStringStringText.Text = "";
                BreakFlowStringStringText.Enabled = false;
                BreakFlowStringIsRegExCheck.Checked = false;
                BreakFlowStringIsRegExCheck.Enabled = false;
                BreakFlowStringFlowPreviousCheck.Checked = false;
                BreakFlowStringFlowPreviousCheck.Enabled = false;
                BreakFlowStringFlowNextCheck.Checked = false;
                BreakFlowStringFlowNextCheck.Enabled = false;
            }
            else
            {
                BreakFlowStringDeleteButton.Enabled = true;
                BreakFlowStringUpButton.Enabled = (index > 0);
                BreakFlowStringDownButton.Enabled = (index < (BulletList.Items.Count-1));

                BreakFlowString lb = (BreakFlowString)mpset.mBreakFlowStrings[index];
                BreakFlowStringStringText.Enabled = true;
                BreakFlowStringStringText.Text = convertToDisplay(lb.mString,lb.mIsRegEx);
                BreakFlowStringIsRegExCheck.Enabled = true;
                BreakFlowStringIsRegExCheck.Checked = lb.mIsRegEx;

                BreakFlowStringFlowPreviousCheck.Checked = lb.mNeverReflowLine;
                BreakFlowStringFlowPreviousCheck.Enabled = true;
                if (BreakFlowStringFlowPreviousCheck.Checked)
                {
                    BreakFlowStringFlowNextCheck.Enabled = false;
                    BreakFlowStringFlowNextCheck.Checked = false;
                }
                else
                {
                    BreakFlowStringFlowNextCheck.Checked = lb.mNeverReflowIntoNextLine;
                    BreakFlowStringFlowNextCheck.Enabled = true;
                }
            }
        }
    
        
        private void selectBreakFlowStringListItem(int index)
        {
            mCurrentBreakFlowStringListItemSelected = index;
            if (index != -1)
            {
                mProgramaticallySettingSelectedBreakFlowString = true;
                BreakFlowStringList.Items[index].Selected = true;
                mProgramaticallySettingSelectedBreakFlowString = false;
            }
        }

        private bool validateSelectedBreakFlowString(bool printError)
        {
            int index = mCurrentBreakFlowStringListItemSelected;
            if (index != -1)
            {
                BreakFlowString bp = (BreakFlowString)mpset.mBreakFlowStrings[index];
                bp.mName = BreakFlowStringList.Items[index].Text;
                bp.mIsRegEx = BreakFlowStringIsRegExCheck.Checked;
                bp.mString = convertFromDisplay(BreakFlowStringStringText.Text,bp.mIsRegEx);
                bp.mNeverReflowLine = BreakFlowStringFlowPreviousCheck.Checked;
                bp.mNeverReflowIntoNextLine = BreakFlowStringFlowNextCheck.Checked;
                try
                {
                    mpset.validateBreakFlowString(index);
                }
                catch (Exception e)
                {
                    if (printError)
                    {
                        MessageBox.Show("Error detected validating BreakFlowString - " + e.Message);
                    }
                    return false;
                }
            }
            return true;
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(CommentReflowerSetup));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.GeneralTab = new System.Windows.Forms.TabPage();
            this.AboutBtn = new System.Windows.Forms.Button();
            this.BlockMinimumWidthText = new System.Windows.Forms.TextBox();
            this.BlockWrapWidthText = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.UseTabsToIndentCheck = new System.Windows.Forms.CheckBox();
            this.BlocksTab = new System.Windows.Forms.TabPage();
            this.FirstLineBlockCombo = new System.Windows.Forms.ComboBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.BlockEndTypeCombo = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.BlockStartTypeCombo = new System.Windows.Forms.ComboBox();
            this.IsBlockLineStartRegExCheck = new System.Windows.Forms.CheckBox();
            this.IsBlockEndRegExCheck = new System.Windows.Forms.CheckBox();
            this.IsBlockStartRegExCheck = new System.Windows.Forms.CheckBox();
            this.BlockLineStartText = new System.Windows.Forms.TextBox();
            this.BlockEndText = new System.Windows.Forms.TextBox();
            this.BlockStartText = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.BlockFileTypeCombo = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.BlockNewButton = new System.Windows.Forms.Button();
            this.BlockUpButton = new System.Windows.Forms.Button();
            this.BlockDownButton = new System.Windows.Forms.Button();
            this.BlockDeleteButton = new System.Windows.Forms.Button();
            this.BlockList = new System.Windows.Forms.ListView();
            this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
            this.BulletsTab = new System.Windows.Forms.TabPage();
            this.BulletEdgeCombo = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.BulletIsRegExCheck = new System.Windows.Forms.CheckBox();
            this.BulletStringText = new System.Windows.Forms.TextBox();
            this.BulletNewButton = new System.Windows.Forms.Button();
            this.BulletUpButton = new System.Windows.Forms.Button();
            this.BulletDownButton = new System.Windows.Forms.Button();
            this.BulletDeleteButton = new System.Windows.Forms.Button();
            this.BulletList = new System.Windows.Forms.ListView();
            this.columnHeader2 = new System.Windows.Forms.ColumnHeader();
            this.BreakFlowStringsTab = new System.Windows.Forms.TabPage();
            this.BreakFlowStringFlowNextCheck = new System.Windows.Forms.CheckBox();
            this.BreakFlowStringFlowPreviousCheck = new System.Windows.Forms.CheckBox();
            this.label11 = new System.Windows.Forms.Label();
            this.BreakFlowStringIsRegExCheck = new System.Windows.Forms.CheckBox();
            this.BreakFlowStringStringText = new System.Windows.Forms.TextBox();
            this.BreakFlowStringNewButton = new System.Windows.Forms.Button();
            this.BreakFlowStringUpButton = new System.Windows.Forms.Button();
            this.BreakFlowStringDownButton = new System.Windows.Forms.Button();
            this.BreakFlowStringDeleteButton = new System.Windows.Forms.Button();
            this.BreakFlowStringList = new System.Windows.Forms.ListView();
            this.columnHeader3 = new System.Windows.Forms.ColumnHeader();
            this.OkBtn = new System.Windows.Forms.Button();
            this.CancelBtn = new System.Windows.Forms.Button();
            this.HelpBtn = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.GeneralTab.SuspendLayout();
            this.BlocksTab.SuspendLayout();
            this.BulletsTab.SuspendLayout();
            this.BreakFlowStringsTab.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.GeneralTab);
            this.tabControl1.Controls.Add(this.BlocksTab);
            this.tabControl1.Controls.Add(this.BulletsTab);
            this.tabControl1.Controls.Add(this.BreakFlowStringsTab);
            this.tabControl1.Location = new System.Drawing.Point(8, 8);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(376, 384);
            this.tabControl1.TabIndex = 0;
            // 
            // GeneralTab
            // 
            this.GeneralTab.Controls.Add(this.AboutBtn);
            this.GeneralTab.Controls.Add(this.BlockMinimumWidthText);
            this.GeneralTab.Controls.Add(this.BlockWrapWidthText);
            this.GeneralTab.Controls.Add(this.label6);
            this.GeneralTab.Controls.Add(this.label5);
            this.GeneralTab.Controls.Add(this.UseTabsToIndentCheck);
            this.GeneralTab.Location = new System.Drawing.Point(4, 22);
            this.GeneralTab.Name = "GeneralTab";
            this.GeneralTab.Size = new System.Drawing.Size(368, 358);
            this.GeneralTab.TabIndex = 0;
            this.GeneralTab.Text = "General";
            this.GeneralTab.Validating += new System.ComponentModel.CancelEventHandler(this.GeneralTab_Validating);
            // 
            // AboutBtn
            // 
            this.AboutBtn.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.AboutBtn.Location = new System.Drawing.Point(288, 328);
            this.AboutBtn.Name = "AboutBtn";
            this.AboutBtn.Size = new System.Drawing.Size(72, 24);
            this.AboutBtn.TabIndex = 0;
            this.AboutBtn.Text = "About";
            this.AboutBtn.Click += new System.EventHandler(this.AboutBtn_Click);
            // 
            // BlockMinimumWidthText
            // 
            this.BlockMinimumWidthText.Location = new System.Drawing.Point(296, 64);
            this.BlockMinimumWidthText.MaxLength = 4;
            this.BlockMinimumWidthText.Name = "BlockMinimumWidthText";
            this.BlockMinimumWidthText.Size = new System.Drawing.Size(56, 20);
            this.BlockMinimumWidthText.TabIndex = 5;
            this.BlockMinimumWidthText.Text = "";
            // 
            // BlockWrapWidthText
            // 
            this.BlockWrapWidthText.Location = new System.Drawing.Point(296, 40);
            this.BlockWrapWidthText.MaxLength = 4;
            this.BlockWrapWidthText.Name = "BlockWrapWidthText";
            this.BlockWrapWidthText.Size = new System.Drawing.Size(56, 20);
            this.BlockWrapWidthText.TabIndex = 3;
            this.BlockWrapWidthText.Text = "";
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(16, 68);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(120, 16);
            this.label6.TabIndex = 4;
            this.label6.Text = "Minimum Block Width:";
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(16, 44);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(72, 16);
            this.label5.TabIndex = 2;
            this.label5.Text = "Wrap Width:";
            // 
            // UseTabsToIndentCheck
            // 
            this.UseTabsToIndentCheck.Checked = true;
            this.UseTabsToIndentCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.UseTabsToIndentCheck.Location = new System.Drawing.Point(24, 16);
            this.UseTabsToIndentCheck.Name = "UseTabsToIndentCheck";
            this.UseTabsToIndentCheck.Size = new System.Drawing.Size(272, 16);
            this.UseTabsToIndentCheck.TabIndex = 1;
            this.UseTabsToIndentCheck.Text = "Use tabs to indent";
            // 
            // BlocksTab
            // 
            this.BlocksTab.Controls.Add(this.FirstLineBlockCombo);
            this.BlocksTab.Controls.Add(this.label12);
            this.BlocksTab.Controls.Add(this.label8);
            this.BlocksTab.Controls.Add(this.BlockEndTypeCombo);
            this.BlocksTab.Controls.Add(this.label7);
            this.BlocksTab.Controls.Add(this.BlockStartTypeCombo);
            this.BlocksTab.Controls.Add(this.IsBlockLineStartRegExCheck);
            this.BlocksTab.Controls.Add(this.IsBlockEndRegExCheck);
            this.BlocksTab.Controls.Add(this.IsBlockStartRegExCheck);
            this.BlocksTab.Controls.Add(this.BlockLineStartText);
            this.BlocksTab.Controls.Add(this.BlockEndText);
            this.BlocksTab.Controls.Add(this.BlockStartText);
            this.BlocksTab.Controls.Add(this.label4);
            this.BlocksTab.Controls.Add(this.label3);
            this.BlocksTab.Controls.Add(this.label2);
            this.BlocksTab.Controls.Add(this.BlockFileTypeCombo);
            this.BlocksTab.Controls.Add(this.label1);
            this.BlocksTab.Controls.Add(this.BlockNewButton);
            this.BlocksTab.Controls.Add(this.BlockUpButton);
            this.BlocksTab.Controls.Add(this.BlockDownButton);
            this.BlocksTab.Controls.Add(this.BlockDeleteButton);
            this.BlocksTab.Controls.Add(this.BlockList);
            this.BlocksTab.Location = new System.Drawing.Point(4, 22);
            this.BlocksTab.Name = "BlocksTab";
            this.BlocksTab.Size = new System.Drawing.Size(368, 358);
            this.BlocksTab.TabIndex = 1;
            this.BlocksTab.Text = "Blocks";
            this.BlocksTab.Validating += new System.ComponentModel.CancelEventHandler(this.BlocksTab_Validating);
            // 
            // FirstLineBlockCombo
            // 
            this.FirstLineBlockCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.FirstLineBlockCombo.Items.AddRange(new object[] {
                                                                     "must not contain text before start",
                                                                     "can contain text before start"});
            this.FirstLineBlockCombo.Location = new System.Drawing.Point(112, 320);
            this.FirstLineBlockCombo.Name = "FirstLineBlockCombo";
            this.FirstLineBlockCombo.Size = new System.Drawing.Size(168, 21);
            this.FirstLineBlockCombo.TabIndex = 21;
            // 
            // label12
            // 
            this.label12.Location = new System.Drawing.Point(8, 325);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(104, 16);
            this.label12.TabIndex = 20;
            this.label12.Text = "First Line of Block:";
            // 
            // label8
            // 
            this.label8.Location = new System.Drawing.Point(8, 253);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(72, 16);
            this.label8.TabIndex = 12;
            this.label8.Text = "Block End is:";
            // 
            // BlockEndTypeCombo
            // 
            this.BlockEndTypeCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.BlockEndTypeCombo.Items.AddRange(new object[] {
                                                                   "not used",
                                                                   "always on own line",
                                                                   "on own line if comment is >1 lines",
                                                                   "never on own line"});
            this.BlockEndTypeCombo.Location = new System.Drawing.Point(112, 248);
            this.BlockEndTypeCombo.Name = "BlockEndTypeCombo";
            this.BlockEndTypeCombo.Size = new System.Drawing.Size(168, 21);
            this.BlockEndTypeCombo.TabIndex = 13;
            this.BlockEndTypeCombo.SelectedIndexChanged += new System.EventHandler(this.BlockEndTypeCombo_SelectedIndexChanged);
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(8, 205);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(80, 16);
            this.label7.TabIndex = 7;
            this.label7.Text = "Block Start is:";
            // 
            // BlockStartTypeCombo
            // 
            this.BlockStartTypeCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.BlockStartTypeCombo.Items.AddRange(new object[] {
                                                                     "not used",
                                                                     "always on own line",
                                                                     "on own line if comment is >1 lines",
                                                                     "never on own line"});
            this.BlockStartTypeCombo.Location = new System.Drawing.Point(112, 200);
            this.BlockStartTypeCombo.Name = "BlockStartTypeCombo";
            this.BlockStartTypeCombo.Size = new System.Drawing.Size(168, 21);
            this.BlockStartTypeCombo.TabIndex = 8;
            this.BlockStartTypeCombo.SelectedIndexChanged += new System.EventHandler(this.BlockStartTypeCombo_SelectedIndexChanged);
            // 
            // IsBlockLineStartRegExCheck
            // 
            this.IsBlockLineStartRegExCheck.Enabled = false;
            this.IsBlockLineStartRegExCheck.Location = new System.Drawing.Point(288, 292);
            this.IsBlockLineStartRegExCheck.Name = "IsBlockLineStartRegExCheck";
            this.IsBlockLineStartRegExCheck.Size = new System.Drawing.Size(72, 24);
            this.IsBlockLineStartRegExCheck.TabIndex = 19;
            this.IsBlockLineStartRegExCheck.Text = "is RegEx";
            // 
            // IsBlockEndRegExCheck
            // 
            this.IsBlockEndRegExCheck.Location = new System.Drawing.Point(288, 268);
            this.IsBlockEndRegExCheck.Name = "IsBlockEndRegExCheck";
            this.IsBlockEndRegExCheck.Size = new System.Drawing.Size(72, 24);
            this.IsBlockEndRegExCheck.TabIndex = 16;
            this.IsBlockEndRegExCheck.Text = "is RegEx";
            // 
            // IsBlockStartRegExCheck
            // 
            this.IsBlockStartRegExCheck.Location = new System.Drawing.Point(288, 220);
            this.IsBlockStartRegExCheck.Name = "IsBlockStartRegExCheck";
            this.IsBlockStartRegExCheck.Size = new System.Drawing.Size(72, 24);
            this.IsBlockStartRegExCheck.TabIndex = 11;
            this.IsBlockStartRegExCheck.Text = "is RegEx";
            // 
            // BlockLineStartText
            // 
            this.BlockLineStartText.Location = new System.Drawing.Point(112, 296);
            this.BlockLineStartText.Name = "BlockLineStartText";
            this.BlockLineStartText.Size = new System.Drawing.Size(168, 20);
            this.BlockLineStartText.TabIndex = 18;
            this.BlockLineStartText.Text = "";
            // 
            // BlockEndText
            // 
            this.BlockEndText.Location = new System.Drawing.Point(112, 272);
            this.BlockEndText.Name = "BlockEndText";
            this.BlockEndText.Size = new System.Drawing.Size(168, 20);
            this.BlockEndText.TabIndex = 15;
            this.BlockEndText.Text = "";
            // 
            // BlockStartText
            // 
            this.BlockStartText.Location = new System.Drawing.Point(112, 224);
            this.BlockStartText.Name = "BlockStartText";
            this.BlockStartText.Size = new System.Drawing.Size(168, 20);
            this.BlockStartText.TabIndex = 10;
            this.BlockStartText.Text = "";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(8, 300);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(72, 16);
            this.label4.TabIndex = 17;
            this.label4.Text = "Line Start:";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(8, 276);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(72, 16);
            this.label3.TabIndex = 14;
            this.label3.Text = "Block End:";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(8, 228);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(72, 16);
            this.label2.TabIndex = 9;
            this.label2.Text = "Block Start:";
            // 
            // BlockFileTypeCombo
            // 
            this.BlockFileTypeCombo.Location = new System.Drawing.Point(112, 176);
            this.BlockFileTypeCombo.Name = "BlockFileTypeCombo";
            this.BlockFileTypeCombo.Size = new System.Drawing.Size(168, 21);
            this.BlockFileTypeCombo.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(8, 181);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 16);
            this.label1.TabIndex = 5;
            this.label1.Text = "File Types:";
            // 
            // BlockNewButton
            // 
            this.BlockNewButton.Image = ((System.Drawing.Image)(resources.GetObject("BlockNewButton.Image")));
            this.BlockNewButton.Location = new System.Drawing.Point(264, 8);
            this.BlockNewButton.Name = "BlockNewButton";
            this.BlockNewButton.Size = new System.Drawing.Size(22, 22);
            this.BlockNewButton.TabIndex = 0;
            this.BlockNewButton.Click += new System.EventHandler(this.BlockNewButton_Click);
            // 
            // BlockUpButton
            // 
            this.BlockUpButton.Image = ((System.Drawing.Image)(resources.GetObject("BlockUpButton.Image")));
            this.BlockUpButton.Location = new System.Drawing.Point(336, 8);
            this.BlockUpButton.Name = "BlockUpButton";
            this.BlockUpButton.Size = new System.Drawing.Size(22, 22);
            this.BlockUpButton.TabIndex = 3;
            this.BlockUpButton.Click += new System.EventHandler(this.BlockUpButton_Click);
            // 
            // BlockDownButton
            // 
            this.BlockDownButton.Image = ((System.Drawing.Image)(resources.GetObject("BlockDownButton.Image")));
            this.BlockDownButton.Location = new System.Drawing.Point(312, 8);
            this.BlockDownButton.Name = "BlockDownButton";
            this.BlockDownButton.Size = new System.Drawing.Size(22, 22);
            this.BlockDownButton.TabIndex = 2;
            this.BlockDownButton.Click += new System.EventHandler(this.BlockDownButton_Click);
            // 
            // BlockDeleteButton
            // 
            this.BlockDeleteButton.Image = ((System.Drawing.Image)(resources.GetObject("BlockDeleteButton.Image")));
            this.BlockDeleteButton.Location = new System.Drawing.Point(288, 8);
            this.BlockDeleteButton.Name = "BlockDeleteButton";
            this.BlockDeleteButton.Size = new System.Drawing.Size(22, 22);
            this.BlockDeleteButton.TabIndex = 1;
            this.BlockDeleteButton.Click += new System.EventHandler(this.BlockDeleteButton_Click);
            // 
            // BlockList
            // 
            this.BlockList.AutoArrange = false;
            this.BlockList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
                                                                                        this.columnHeader1});
            this.BlockList.FullRowSelect = true;
            this.BlockList.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.BlockList.HideSelection = false;
            this.BlockList.LabelEdit = true;
            this.BlockList.LabelWrap = false;
            this.BlockList.Location = new System.Drawing.Point(8, 32);
            this.BlockList.MultiSelect = false;
            this.BlockList.Name = "BlockList";
            this.BlockList.Size = new System.Drawing.Size(352, 136);
            this.BlockList.TabIndex = 4;
            this.BlockList.View = System.Windows.Forms.View.Details;
            this.BlockList.MouseUp += new System.Windows.Forms.MouseEventHandler(this.BlockList_MouseUp);
            this.BlockList.SelectedIndexChanged += new System.EventHandler(this.BlockList_SelectedIndexChanged);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "";
            this.columnHeader1.Width = 1000;
            // 
            // BulletsTab
            // 
            this.BulletsTab.Controls.Add(this.BulletEdgeCombo);
            this.BulletsTab.Controls.Add(this.label10);
            this.BulletsTab.Controls.Add(this.label9);
            this.BulletsTab.Controls.Add(this.BulletIsRegExCheck);
            this.BulletsTab.Controls.Add(this.BulletStringText);
            this.BulletsTab.Controls.Add(this.BulletNewButton);
            this.BulletsTab.Controls.Add(this.BulletUpButton);
            this.BulletsTab.Controls.Add(this.BulletDownButton);
            this.BulletsTab.Controls.Add(this.BulletDeleteButton);
            this.BulletsTab.Controls.Add(this.BulletList);
            this.BulletsTab.Location = new System.Drawing.Point(4, 22);
            this.BulletsTab.Name = "BulletsTab";
            this.BulletsTab.Size = new System.Drawing.Size(368, 358);
            this.BulletsTab.TabIndex = 2;
            this.BulletsTab.Text = "Bullets";
            this.BulletsTab.Validating += new System.ComponentModel.CancelEventHandler(this.BulletsTab_Validating);
            // 
            // BulletEdgeCombo
            // 
            this.BulletEdgeCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.BulletEdgeCombo.Items.AddRange(new object[] {
                                                                 "Left",
                                                                 "Right"});
            this.BulletEdgeCombo.Location = new System.Drawing.Point(192, 200);
            this.BulletEdgeCombo.Name = "BulletEdgeCombo";
            this.BulletEdgeCombo.Size = new System.Drawing.Size(88, 21);
            this.BulletEdgeCombo.TabIndex = 9;
            // 
            // label10
            // 
            this.label10.Location = new System.Drawing.Point(8, 205);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(80, 16);
            this.label10.TabIndex = 8;
            this.label10.Text = "Wrap on Edge: ";
            // 
            // label9
            // 
            this.label9.Location = new System.Drawing.Point(8, 180);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(72, 16);
            this.label9.TabIndex = 5;
            this.label9.Text = "Match String:";
            // 
            // BulletIsRegExCheck
            // 
            this.BulletIsRegExCheck.Location = new System.Drawing.Point(288, 172);
            this.BulletIsRegExCheck.Name = "BulletIsRegExCheck";
            this.BulletIsRegExCheck.Size = new System.Drawing.Size(72, 24);
            this.BulletIsRegExCheck.TabIndex = 7;
            this.BulletIsRegExCheck.Text = "is RegEx";
            // 
            // BulletStringText
            // 
            this.BulletStringText.Location = new System.Drawing.Point(112, 176);
            this.BulletStringText.Name = "BulletStringText";
            this.BulletStringText.Size = new System.Drawing.Size(168, 20);
            this.BulletStringText.TabIndex = 6;
            this.BulletStringText.Text = "";
            // 
            // BulletNewButton
            // 
            this.BulletNewButton.Image = ((System.Drawing.Image)(resources.GetObject("BulletNewButton.Image")));
            this.BulletNewButton.Location = new System.Drawing.Point(264, 8);
            this.BulletNewButton.Name = "BulletNewButton";
            this.BulletNewButton.Size = new System.Drawing.Size(22, 22);
            this.BulletNewButton.TabIndex = 0;
            this.BulletNewButton.Click += new System.EventHandler(this.BulletNewButton_Click);
            // 
            // BulletUpButton
            // 
            this.BulletUpButton.Image = ((System.Drawing.Image)(resources.GetObject("BulletUpButton.Image")));
            this.BulletUpButton.Location = new System.Drawing.Point(336, 8);
            this.BulletUpButton.Name = "BulletUpButton";
            this.BulletUpButton.Size = new System.Drawing.Size(22, 22);
            this.BulletUpButton.TabIndex = 3;
            this.BulletUpButton.Click += new System.EventHandler(this.BulletUpButton_Click);
            // 
            // BulletDownButton
            // 
            this.BulletDownButton.Image = ((System.Drawing.Image)(resources.GetObject("BulletDownButton.Image")));
            this.BulletDownButton.Location = new System.Drawing.Point(312, 8);
            this.BulletDownButton.Name = "BulletDownButton";
            this.BulletDownButton.Size = new System.Drawing.Size(22, 22);
            this.BulletDownButton.TabIndex = 2;
            this.BulletDownButton.Click += new System.EventHandler(this.BulletDownButton_Click);
            // 
            // BulletDeleteButton
            // 
            this.BulletDeleteButton.Image = ((System.Drawing.Image)(resources.GetObject("BulletDeleteButton.Image")));
            this.BulletDeleteButton.Location = new System.Drawing.Point(288, 8);
            this.BulletDeleteButton.Name = "BulletDeleteButton";
            this.BulletDeleteButton.Size = new System.Drawing.Size(22, 22);
            this.BulletDeleteButton.TabIndex = 1;
            this.BulletDeleteButton.Click += new System.EventHandler(this.BulletDeleteButton_Click);
            // 
            // BulletList
            // 
            this.BulletList.AutoArrange = false;
            this.BulletList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
                                                                                         this.columnHeader2});
            this.BulletList.FullRowSelect = true;
            this.BulletList.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.BulletList.HideSelection = false;
            this.BulletList.LabelEdit = true;
            this.BulletList.Location = new System.Drawing.Point(8, 32);
            this.BulletList.MultiSelect = false;
            this.BulletList.Name = "BulletList";
            this.BulletList.Size = new System.Drawing.Size(352, 136);
            this.BulletList.TabIndex = 4;
            this.BulletList.View = System.Windows.Forms.View.Details;
            this.BulletList.MouseUp += new System.Windows.Forms.MouseEventHandler(this.BulletList_MouseUp);
            this.BulletList.SelectedIndexChanged += new System.EventHandler(this.BulletList_SelectedIndexChanged);
            // 
            // columnHeader2
            // 
            this.columnHeader2.Width = 1000;
            // 
            // BreakFlowStringsTab
            // 
            this.BreakFlowStringsTab.Controls.Add(this.BreakFlowStringFlowNextCheck);
            this.BreakFlowStringsTab.Controls.Add(this.BreakFlowStringFlowPreviousCheck);
            this.BreakFlowStringsTab.Controls.Add(this.label11);
            this.BreakFlowStringsTab.Controls.Add(this.BreakFlowStringIsRegExCheck);
            this.BreakFlowStringsTab.Controls.Add(this.BreakFlowStringStringText);
            this.BreakFlowStringsTab.Controls.Add(this.BreakFlowStringNewButton);
            this.BreakFlowStringsTab.Controls.Add(this.BreakFlowStringUpButton);
            this.BreakFlowStringsTab.Controls.Add(this.BreakFlowStringDownButton);
            this.BreakFlowStringsTab.Controls.Add(this.BreakFlowStringDeleteButton);
            this.BreakFlowStringsTab.Controls.Add(this.BreakFlowStringList);
            this.BreakFlowStringsTab.Location = new System.Drawing.Point(4, 22);
            this.BreakFlowStringsTab.Name = "BreakFlowStringsTab";
            this.BreakFlowStringsTab.Size = new System.Drawing.Size(368, 358);
            this.BreakFlowStringsTab.TabIndex = 3;
            this.BreakFlowStringsTab.Text = "Break Flow Strings";
            this.BreakFlowStringsTab.Validating += new System.ComponentModel.CancelEventHandler(this.BreakFlowStringsTab_Validating);
            // 
            // BreakFlowStringFlowNextCheck
            // 
            this.BreakFlowStringFlowNextCheck.Checked = true;
            this.BreakFlowStringFlowNextCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.BreakFlowStringFlowNextCheck.Location = new System.Drawing.Point(16, 232);
            this.BreakFlowStringFlowNextCheck.Name = "BreakFlowStringFlowNextCheck";
            this.BreakFlowStringFlowNextCheck.Size = new System.Drawing.Size(272, 16);
            this.BreakFlowStringFlowNextCheck.TabIndex = 9;
            this.BreakFlowStringFlowNextCheck.Text = "Never reflow line with this string into next";
            // 
            // BreakFlowStringFlowPreviousCheck
            // 
            this.BreakFlowStringFlowPreviousCheck.Checked = true;
            this.BreakFlowStringFlowPreviousCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.BreakFlowStringFlowPreviousCheck.Location = new System.Drawing.Point(16, 208);
            this.BreakFlowStringFlowPreviousCheck.Name = "BreakFlowStringFlowPreviousCheck";
            this.BreakFlowStringFlowPreviousCheck.Size = new System.Drawing.Size(272, 16);
            this.BreakFlowStringFlowPreviousCheck.TabIndex = 8;
            this.BreakFlowStringFlowPreviousCheck.Text = "Never reflow line with this string at all";
            this.BreakFlowStringFlowPreviousCheck.CheckedChanged += new System.EventHandler(this.BreakFlowStringFlowPreviousCheck_CheckedChanged);
            // 
            // label11
            // 
            this.label11.Location = new System.Drawing.Point(8, 180);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(72, 16);
            this.label11.TabIndex = 5;
            this.label11.Text = "Match String:";
            // 
            // BreakFlowStringIsRegExCheck
            // 
            this.BreakFlowStringIsRegExCheck.Location = new System.Drawing.Point(288, 172);
            this.BreakFlowStringIsRegExCheck.Name = "BreakFlowStringIsRegExCheck";
            this.BreakFlowStringIsRegExCheck.Size = new System.Drawing.Size(72, 24);
            this.BreakFlowStringIsRegExCheck.TabIndex = 7;
            this.BreakFlowStringIsRegExCheck.Text = "is RegEx";
            // 
            // BreakFlowStringStringText
            // 
            this.BreakFlowStringStringText.Location = new System.Drawing.Point(112, 176);
            this.BreakFlowStringStringText.Name = "BreakFlowStringStringText";
            this.BreakFlowStringStringText.Size = new System.Drawing.Size(168, 20);
            this.BreakFlowStringStringText.TabIndex = 6;
            this.BreakFlowStringStringText.Text = "";
            // 
            // BreakFlowStringNewButton
            // 
            this.BreakFlowStringNewButton.Image = ((System.Drawing.Image)(resources.GetObject("BreakFlowStringNewButton.Image")));
            this.BreakFlowStringNewButton.Location = new System.Drawing.Point(264, 8);
            this.BreakFlowStringNewButton.Name = "BreakFlowStringNewButton";
            this.BreakFlowStringNewButton.Size = new System.Drawing.Size(22, 22);
            this.BreakFlowStringNewButton.TabIndex = 0;
            this.BreakFlowStringNewButton.Click += new System.EventHandler(this.BreakFlowStringNewButton_Click);
            // 
            // BreakFlowStringUpButton
            // 
            this.BreakFlowStringUpButton.Image = ((System.Drawing.Image)(resources.GetObject("BreakFlowStringUpButton.Image")));
            this.BreakFlowStringUpButton.Location = new System.Drawing.Point(336, 8);
            this.BreakFlowStringUpButton.Name = "BreakFlowStringUpButton";
            this.BreakFlowStringUpButton.Size = new System.Drawing.Size(22, 22);
            this.BreakFlowStringUpButton.TabIndex = 3;
            this.BreakFlowStringUpButton.Click += new System.EventHandler(this.BreakFlowStringUpButton_Click);
            // 
            // BreakFlowStringDownButton
            // 
            this.BreakFlowStringDownButton.Image = ((System.Drawing.Image)(resources.GetObject("BreakFlowStringDownButton.Image")));
            this.BreakFlowStringDownButton.Location = new System.Drawing.Point(312, 8);
            this.BreakFlowStringDownButton.Name = "BreakFlowStringDownButton";
            this.BreakFlowStringDownButton.Size = new System.Drawing.Size(22, 22);
            this.BreakFlowStringDownButton.TabIndex = 2;
            this.BreakFlowStringDownButton.Click += new System.EventHandler(this.BreakFlowStringDownButton_Click);
            // 
            // BreakFlowStringDeleteButton
            // 
            this.BreakFlowStringDeleteButton.Image = ((System.Drawing.Image)(resources.GetObject("BreakFlowStringDeleteButton.Image")));
            this.BreakFlowStringDeleteButton.Location = new System.Drawing.Point(288, 8);
            this.BreakFlowStringDeleteButton.Name = "BreakFlowStringDeleteButton";
            this.BreakFlowStringDeleteButton.Size = new System.Drawing.Size(22, 22);
            this.BreakFlowStringDeleteButton.TabIndex = 1;
            this.BreakFlowStringDeleteButton.Click += new System.EventHandler(this.BreakFlowStringDeleteButton_Click);
            // 
            // BreakFlowStringList
            // 
            this.BreakFlowStringList.AutoArrange = false;
            this.BreakFlowStringList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
                                                                                                  this.columnHeader3});
            this.BreakFlowStringList.FullRowSelect = true;
            this.BreakFlowStringList.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.BreakFlowStringList.HideSelection = false;
            this.BreakFlowStringList.LabelEdit = true;
            this.BreakFlowStringList.Location = new System.Drawing.Point(8, 32);
            this.BreakFlowStringList.MultiSelect = false;
            this.BreakFlowStringList.Name = "BreakFlowStringList";
            this.BreakFlowStringList.Size = new System.Drawing.Size(352, 136);
            this.BreakFlowStringList.TabIndex = 4;
            this.BreakFlowStringList.View = System.Windows.Forms.View.Details;
            this.BreakFlowStringList.MouseUp += new System.Windows.Forms.MouseEventHandler(this.BreakFlowStringList_MouseUp);
            this.BreakFlowStringList.SelectedIndexChanged += new System.EventHandler(this.BreakFlowStringList_SelectedIndexChanged);
            // 
            // columnHeader3
            // 
            this.columnHeader3.Width = 1000;
            // 
            // OkBtn
            // 
            this.OkBtn.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.OkBtn.Location = new System.Drawing.Point(152, 400);
            this.OkBtn.Name = "OkBtn";
            this.OkBtn.Size = new System.Drawing.Size(72, 24);
            this.OkBtn.TabIndex = 1;
            this.OkBtn.Text = "OK";
            // 
            // CancelBtn
            // 
            this.CancelBtn.CausesValidation = false;
            this.CancelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelBtn.Location = new System.Drawing.Point(232, 400);
            this.CancelBtn.Name = "CancelBtn";
            this.CancelBtn.Size = new System.Drawing.Size(72, 24);
            this.CancelBtn.TabIndex = 2;
            this.CancelBtn.Text = "Cancel";
            // 
            // HelpBtn
            // 
            this.HelpBtn.Location = new System.Drawing.Point(312, 400);
            this.HelpBtn.Name = "HelpBtn";
            this.HelpBtn.Size = new System.Drawing.Size(72, 24);
            this.HelpBtn.TabIndex = 3;
            this.HelpBtn.Text = "Help";
            this.HelpBtn.Click += new System.EventHandler(this.HelpBtn_Click);
            // 
            // CommentReflowerSetup
            // 
            this.AcceptButton = this.OkBtn;
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.CancelButton = this.CancelBtn;
            this.ClientSize = new System.Drawing.Size(392, 430);
            this.Controls.Add(this.HelpBtn);
            this.Controls.Add(this.CancelBtn);
            this.Controls.Add(this.OkBtn);
            this.Controls.Add(this.tabControl1);
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(400, 464);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(400, 464);
            this.Name = "CommentReflowerSetup";
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Comment Reflower Setup";
            this.TopMost = true;
            this.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.CommentReflowerSetup_HelpRequested);
            this.MouseUp += new System.Windows.Forms.MouseEventHandler(this.BulletList_MouseUp);
            this.tabControl1.ResumeLayout(false);
            this.GeneralTab.ResumeLayout(false);
            this.BlocksTab.ResumeLayout(false);
            this.BulletsTab.ResumeLayout(false);
            this.BreakFlowStringsTab.ResumeLayout(false);
            this.ResumeLayout(false);

        }
        #endregion

        private void BulletDeleteButton_Click(object sender, System.EventArgs e)
        {
            int index = BulletList.SelectedItems[0].Index;

            BulletList.Items[index].Remove();
            mpset.mBulletPoints.RemoveAt(index);

            if (index >= BulletList.Items.Count)
            {
                index = index - 1;
            }
            selectBulletListItem(index);
            if (index != -1)
            {
                BulletList.Items[index].EnsureVisible();
            }
            updateItemsForBullet(index);//this will be -1 if last item is deleted
        }

        private void BulletNewButton_Click(object sender, System.EventArgs e)
        {
            if (!validateSelectedBullet(true))
            {
                return;
            }
            BulletPoint newObj = new BulletPoint("New Bullet Point", ".*",true, true);
            mpset.mBulletPoints.Add(newObj);
            BulletList.Items.Add(new ListViewItem(newObj.mName));
            int index = BulletList.Items.Count-1;
            selectBulletListItem(index);
            BulletList.Items[index].EnsureVisible();
            BulletList.Items[index].BeginEdit();
            updateItemsForBullet(index);
        }
        private void BulletDownButton_Click(object sender, System.EventArgs e)
        {
            int index = BulletList.SelectedItems[0].Index;

            string temp = BulletList.Items[index].Text;
            BulletList.Items[index].Text = BulletList.Items[index+1].Text;
            BulletList.Items[index+1].Text = temp;

            object tempo = mpset.mBulletPoints[index];
            mpset.mBulletPoints[index] = mpset.mBulletPoints[index+1];
            mpset.mBulletPoints[index+1] = tempo;

            selectBulletListItem(index+1);
            BulletList.Items[index+1].EnsureVisible();
            updateItemsForBullet(index+1);
        }

        private void BulletUpButton_Click(object sender, System.EventArgs e)
        {
            int index = BulletList.SelectedItems[0].Index;

            string temp = BulletList.Items[index].Text;
            BulletList.Items[index].Text = BulletList.Items[index-1].Text;
            BulletList.Items[index-1].Text = temp;

            object tempo = mpset.mBulletPoints[index];
            mpset.mBulletPoints[index] = mpset.mBulletPoints[index-1];
            mpset.mBulletPoints[index-1] = tempo;

            selectBulletListItem(index-1);
            BulletList.Items[index-1].EnsureVisible();
            updateItemsForBullet(index-1);
        }

        private void BulletList_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (!mProgramaticallySettingSelectedBullet)
            {
                // user selected new bullet
                if (BulletList.SelectedItems.Count > 0)
                {
                    int index = BulletList.SelectedItems[0].Index;
                    if (!validateSelectedBullet(mFirstBulletListError))
                    {
                        selectBulletListItem(mCurrentBulletListItemSelected);
                        mIndexToSelectOnBulletListMouseUp = mCurrentBulletListItemSelected;
                        mFirstBulletListError = !mFirstBulletListError;
                    }
                    else
                    {
                        mCurrentBulletListItemSelected = index;
                        updateItemsForBullet(index);
                    }
                }
            }
        }

        private void BulletList_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (mIndexToSelectOnBulletListMouseUp != -1)
            {
                selectBulletListItem(mIndexToSelectOnBulletListMouseUp);
                mIndexToSelectOnBulletListMouseUp = -1;
            }
        }

        private void BulletsTab_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = !validateSelectedBullet(true);
        }

        private void BreakFlowStringDeleteButton_Click(object sender, System.EventArgs e)
        {
            int index = BreakFlowStringList.SelectedItems[0].Index;

            BreakFlowStringList.Items[index].Remove();
            mpset.mBreakFlowStrings.RemoveAt(index);

            if (index >= BreakFlowStringList.Items.Count)
            {
                index = index - 1;
            }
            selectBreakFlowStringListItem(index);
            if (index != -1)
            {
                BreakFlowStringList.Items[index].EnsureVisible();
            }
            updateItemsForBreakFlowString(index);//this will be -1 if last item is deleted
        }

        private void BreakFlowStringNewButton_Click(object sender, System.EventArgs e)
        {
            if (!validateSelectedBreakFlowString(true))
            {
                return;
            }
            BreakFlowString newObj = new BreakFlowString("New break flow string", "string to break flow",false,true,true);
            mpset.mBreakFlowStrings.Add(newObj);
            BreakFlowStringList.Items.Add(new ListViewItem(newObj.mName));
            int index = BreakFlowStringList.Items.Count-1;
            selectBreakFlowStringListItem(index);
            BreakFlowStringList.Items[index].EnsureVisible();
            BreakFlowStringList.Items[index].BeginEdit();
            updateItemsForBreakFlowString(index);
        }

        private void BreakFlowStringDownButton_Click(object sender, System.EventArgs e)
        {
            int index = BreakFlowStringList.SelectedItems[0].Index;

            string temp = BreakFlowStringList.Items[index].Text;
            BreakFlowStringList.Items[index].Text = BreakFlowStringList.Items[index+1].Text;
            BreakFlowStringList.Items[index+1].Text = temp;

            object tempo = mpset.mBreakFlowStrings[index];
            mpset.mBreakFlowStrings[index] = mpset.mBreakFlowStrings[index+1];
            mpset.mBreakFlowStrings[index+1] = tempo;

            selectBreakFlowStringListItem(index+1);
            BreakFlowStringList.Items[index+1].EnsureVisible();
            updateItemsForBreakFlowString(index+1);
        }

        private void BreakFlowStringUpButton_Click(object sender, System.EventArgs e)
        {
            int index = BreakFlowStringList.SelectedItems[0].Index;

            string temp = BreakFlowStringList.Items[index].Text;
            BreakFlowStringList.Items[index].Text = BreakFlowStringList.Items[index-1].Text;
            BreakFlowStringList.Items[index-1].Text = temp;

            object tempo = mpset.mBreakFlowStrings[index];
            mpset.mBreakFlowStrings[index] = mpset.mBreakFlowStrings[index-1];
            mpset.mBreakFlowStrings[index-1] = tempo;

            selectBreakFlowStringListItem(index-1);
            BreakFlowStringList.Items[index-1].EnsureVisible();
            updateItemsForBreakFlowString(index-1);
        }

        private void BreakFlowStringList_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (!mProgramaticallySettingSelectedBreakFlowString)
            {
                // user selected new BreakFlowString
                if (BreakFlowStringList.SelectedItems.Count > 0)
                {
                    int index = BreakFlowStringList.SelectedItems[0].Index;
                    if (!validateSelectedBreakFlowString(mFirstBreakFlowStringListError))
                    {
                        selectBreakFlowStringListItem(mCurrentBreakFlowStringListItemSelected);
                        mIndexToSelectOnBreakFlowStringListMouseUp = mCurrentBreakFlowStringListItemSelected;
                        mFirstBreakFlowStringListError = !mFirstBreakFlowStringListError;
                    }
                    else
                    {
                        mCurrentBreakFlowStringListItemSelected = index;
                        updateItemsForBreakFlowString(index);
                    }
                }
            }
        }

        private void BreakFlowStringList_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (mIndexToSelectOnBreakFlowStringListMouseUp != -1)
            {
                selectBreakFlowStringListItem(mIndexToSelectOnBreakFlowStringListMouseUp);
                mIndexToSelectOnBreakFlowStringListMouseUp = -1;
            }
        }

        private void BreakFlowStringsTab_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = !validateSelectedBreakFlowString(true);
        }


        private void BreakFlowStringFlowPreviousCheck_CheckedChanged(object sender, System.EventArgs e)
        {
            if (BreakFlowStringList.SelectedItems.Count > 0)
            {
                int index = BreakFlowStringList.SelectedItems[0].Index;
                BreakFlowString lb = (BreakFlowString)mpset.mBreakFlowStrings[index];
                if (BreakFlowStringFlowPreviousCheck.Checked)
                {
                    BreakFlowStringFlowNextCheck.Enabled = false;
                    BreakFlowStringFlowNextCheck.Checked = false;
                }
                else
                {
                    BreakFlowStringFlowNextCheck.Checked = lb.mNeverReflowIntoNextLine;
                    BreakFlowStringFlowNextCheck.Enabled = true;
                }
            }
        }


        private void BlockDeleteButton_Click(object sender, System.EventArgs e)
        {
            int index = BlockList.SelectedItems[0].Index;

            BlockList.Items[index].Remove();
            mpset.mCommentBlocks.RemoveAt(index);

            if (index >= BlockList.Items.Count)
            {
                index = index - 1;
            }
            selectBlockListItem(index);
            if (index != -1)
            {
                BlockList.Items[index].EnsureVisible();
            }
            updateItemsForBlock(index);//this will be -1 if last item is deleted
        }

        private void BlockNewButton_Click(object sender, System.EventArgs e)
        {
            if (!validateSelectedBlock(true))
            {
                return;
            }
            ArrayList cPlusPlusCba = new ArrayList();
            cPlusPlusCba.Add("*.new");
            CommentBlock newObj = new CommentBlock(
                "New comment block",
                (ArrayList)cPlusPlusCba.Clone(),
                StartEndBlockType.Empty,
                "",
                false,
                StartEndBlockType.Empty,
                "",
                false,
                "#",
                false);
            mpset.mCommentBlocks.Add(newObj);
            BlockList.Items.Add(new ListViewItem(newObj.mName));
            int index = BlockList.Items.Count-1;
            selectBlockListItem(index);
            BlockList.Items[index].EnsureVisible();
            BlockList.Items[index].BeginEdit();
            updateItemsForBlock(index);
        }
        private void BlockDownButton_Click(object sender, System.EventArgs e)
        {
            int index = BlockList.SelectedItems[0].Index;

            string temp = BlockList.Items[index].Text;
            BlockList.Items[index].Text = BlockList.Items[index+1].Text;
            BlockList.Items[index+1].Text = temp;

            object tempo = mpset.mCommentBlocks[index];
            mpset.mCommentBlocks[index] = mpset.mCommentBlocks[index+1];
            mpset.mCommentBlocks[index+1] = tempo;

            selectBlockListItem(index+1);
            BlockList.Items[index+1].EnsureVisible();
            updateItemsForBlock(index+1);
        }

        private void BlockUpButton_Click(object sender, System.EventArgs e)
        {
            int index = BlockList.SelectedItems[0].Index;

            string temp = BlockList.Items[index].Text;
            BlockList.Items[index].Text = BlockList.Items[index-1].Text;
            BlockList.Items[index-1].Text = temp;

            object tempo = mpset.mCommentBlocks[index];
            mpset.mCommentBlocks[index] = mpset.mCommentBlocks[index-1];
            mpset.mCommentBlocks[index-1] = tempo;

            selectBlockListItem(index-1);
            BlockList.Items[index-1].EnsureVisible();
            updateItemsForBlock(index-1);
        }

        private void BlockList_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (!mProgramaticallySettingSelectedBlock)
            {
                //user selected new Block
                if (BlockList.SelectedItems.Count > 0)
                {
                    int index = BlockList.SelectedItems[0].Index;
                    if (!validateSelectedBlock(mFirstBlockListError))
                    {
                        selectBlockListItem(mCurrentBlockListItemSelected);
                        mIndexToSelectOnBlockListMouseUp = mCurrentBlockListItemSelected;
                        mFirstBlockListError = !mFirstBlockListError;
                    }
                    else
                    {
                        mCurrentBlockListItemSelected = index;
                        updateItemsForBlock(index);
                    }
                }
            }
        }

        private void BlockList_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (mIndexToSelectOnBlockListMouseUp != -1)
            {
                selectBlockListItem(mIndexToSelectOnBlockListMouseUp);
                mIndexToSelectOnBlockListMouseUp = -1;
            }
        }

        private void BlocksTab_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = !validateSelectedBlock(true);
        }

        private void BlockStartTypeCombo_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (BlockStartTypeCombo.SelectedIndex != 0)
            {
                BlockStartText.Enabled = true;
                IsBlockStartRegExCheck.Enabled = true;
            }
            else
            {
                BlockStartText.Enabled = false;
                BlockStartText.Text = "";
                IsBlockStartRegExCheck.Enabled = false;
                IsBlockStartRegExCheck.Checked = false;
            }
        }

        private void BlockEndTypeCombo_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (BlockEndTypeCombo.SelectedIndex != 0)
            {
                BlockEndText.Enabled = true;
                IsBlockEndRegExCheck.Enabled = true;
            }
            else
            {
                BlockEndText.Enabled = false;
                BlockEndText.Text = "";
                IsBlockEndRegExCheck.Enabled = false;
                IsBlockEndRegExCheck.Checked = false;
            }
        }

        private void GeneralTab_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            mpset.mUseTabsToIndent = UseTabsToIndentCheck.Checked;
            try
            {
                mpset.mWrapWidth = System.Convert.ToInt32(BlockWrapWidthText.Text,10);
                mpset.mMinimumBlockWidth = System.Convert.ToInt32(BlockMinimumWidthText.Text,10);
            }
            catch (Exception)
            {
                throw new System.ArgumentException("Wrap width and minumum block width must be integers.");
            }
            mpset.validateGeneralSettings();
        }

        private void AboutBtn_Click(object sender, System.EventArgs e)
        {
            MessageBox.Show(this, "Comment Reflower 1.1\nCopyright (C) 2004 Ian Nowland");
        }

        private void CommentReflowerSetup_HelpRequested(object sender, System.Windows.Forms.HelpEventArgs hlpevent)
        {
            string keyword;
            if (tabControl1.SelectedIndex == 0)
            {
                keyword = "GeneralSettings.htm";
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                keyword = "CommentBlockSettings.htm";
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                keyword = "BulletPointSettings.htm";
            }
            else //if (tabControl1.SelectedIndex == 3)
            {
                keyword = "BreakFlowStringsSettings.htm";
            }

            RegistryKey regkey = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\VisualStudio\7.1\AddIns\CommentReflower.Connect");
            string st1 = (string)regkey.GetValue("SatelliteDllPath");
            Help.ShowHelp(this, st1 + mHelpFile,keyword);
            hlpevent.Handled = true;
        }

        private void HelpBtn_Click(object sender, System.EventArgs e)
        {
            RegistryKey regkey = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\VisualStudio\7.1\AddIns\CommentReflower.Connect");
            string st1 = (string)regkey.GetValue("SatelliteDllPath");
            HelpNavigator navigator = HelpNavigator.TableOfContents;
            Help.ShowHelp(this, st1 + mHelpFile, navigator);
        }


        private string convertToDisplay(string str, bool isRegEx)
        {
            return str.Replace("\t","\\t");
        }

        private string convertFromDisplay(string str, bool isRegEx)
        {
            return str.Replace("\\t","\t");
        }

        private const string mHelpFile = @"CommentReflowerHelp.chm";
    }
}
