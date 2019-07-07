//#define DEBUG_FILE_OUTPUT
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using Advanced_Combat_Tracker;
using System.IO;
using System.Reflection;
using System.Xml;
using System.Text.RegularExpressions;
using System.Threading;
using System.Globalization;

[assembly: AssemblyTitle("Japanese(JPN) Parsing Engine")]
[assembly: AssemblyDescription("Plugin based parsing engine for Japanese EQ2 servers running the Japanese client")]
[assembly: AssemblyCompany("Mayia of Sebilis")]
[assembly: AssemblyVersion("1.1.0.4")]

// NOTE: このpluginは、ACT公式のEn-Jp_ParserのVer.1.1.1.9を元に、日本語環境のままで使用できるように改造したものです。ほぼ全ての機能を取り込んでいると思います。
// NOTE: 解析者向け（＝自分用）に「pluginで解析できなかったログをファイルに出力する」隠し機能を搭載しております。ファイルの１行目の // を外すと利用可能。
// NOTE: Obanburumai様のご協力により、EQ2側のバグによりアーツ名の直後で改行されていて取り込めなかったアーツ（ラッキー・ギャンビット、ワイルド・アクリーション、他にもあるかも？）を取得できるようにしております。（2015/01現在、JPNプラグイン固有の機能です）
// NOTE: 【不具合修正】「マーセナリー・スタンド」が、アーツではなくキャラクターとしてカウントされてしまう現象を修正しました。
////////////////////////////////////////////////////////////////////////////////
// $Date: 2015-10-16 15:27:42 +0900 (2015/10/16 (金)) $
// $Rev: 31 $
////////////////////////////////////////////////////////////////////////////////
namespace ACT_Plugin
{
    public class ACT_Jpn_Parser : UserControl, IActPluginV1
    {
        #region Designer generated code (Avoid editing)
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
            this.cbMultiDamageIsOne = new System.Windows.Forms.CheckBox();
            this.cbRecalcWardedHits = new System.Windows.Forms.CheckBox();
            this.cbKatakana = new System.Windows.Forms.CheckBox();
            this.cbSParseConsider = new System.Windows.Forms.CheckBox();
            this.cbIncludeInterceptFocus = new System.Windows.Forms.CheckBox();
            #if DEBUG_FILE_OUTPUT
            this.cbDebugLog = new System.Windows.Forms.CheckBox();
            this.tbDebugFileName = new System.Windows.Forms.TextBox();
            this.btDebugFileName = new System.Windows.Forms.Button();
            #endif
            this.SuspendLayout();
            // 
            // cbMultiDamageIsOne
            // 
            this.cbMultiDamageIsOne.AutoSize = true;
            this.cbMultiDamageIsOne.Checked = true;
            this.cbMultiDamageIsOne.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbMultiDamageIsOne.Location = new System.Drawing.Point(13, 14);
            this.cbMultiDamageIsOne.Margin = new System.Windows.Forms.Padding(3, 1, 3, 1);
            this.cbMultiDamageIsOne.Name = "cbMultiDamageIsOne";
            this.cbMultiDamageIsOne.Size = new System.Drawing.Size(362, 17);
            this.cbMultiDamageIsOne.TabIndex = 5;
            this.cbMultiDamageIsOne.Text = "複数属性ダメージを1回攻撃として記録します。(既に読み込んだデータには反映しません)";
            this.cbMultiDamageIsOne.MouseHover += new System.EventHandler(this.cbMultiDamageIsOne_MouseHover);
            // 
            // cbRecalcWardedHits
            // 
            this.cbRecalcWardedHits.AutoSize = true;
            this.cbRecalcWardedHits.Checked = true;
            this.cbRecalcWardedHits.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbRecalcWardedHits.Location = new System.Drawing.Point(13, 52);
            this.cbRecalcWardedHits.Margin = new System.Windows.Forms.Padding(3, 1, 3, 1);
            this.cbRecalcWardedHits.Name = "cbRecalcWardedHits";
            this.cbRecalcWardedHits.Size = new System.Drawing.Size(368, 17);
            this.cbRecalcWardedHits.TabIndex = 7;
            this.cbRecalcWardedHits.Text = "Ward で受けた値を本来の値で再計算します。(既に読み込んだデータには反映しません)";
            this.cbRecalcWardedHits.MouseHover += new System.EventHandler(this.cbRecalcWardedHits_MouseHover);
            // 
            // cbKatakana
            // 
            this.cbKatakana.AutoSize = true;
            this.cbKatakana.Checked = true;
            this.cbKatakana.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbKatakana.Location = new System.Drawing.Point(13, 33);
            this.cbKatakana.Margin = new System.Windows.Forms.Padding(3, 1, 3, 1);
            this.cbKatakana.Name = "cbKatakana";
            this.cbKatakana.Size = new System.Drawing.Size(403, 17);
            this.cbKatakana.TabIndex = 9;
            this.cbKatakana.Text = "スキル名の表記を日本語にします。(既に読み込んだデータには反映しません)";
            this.cbKatakana.MouseHover += new System.EventHandler(this.cbKatakana_MouseHover);
            // 
            // cbSParseConsider
            // 
            this.cbSParseConsider.AutoSize = true;
            this.cbSParseConsider.Checked = true;
            this.cbSParseConsider.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbSParseConsider.Location = new System.Drawing.Point(13, 71);
            this.cbSParseConsider.Margin = new System.Windows.Forms.Padding(3, 1, 3, 1);
            this.cbSParseConsider.Name = "cbSParseConsider";
            this.cbSParseConsider.Size = new System.Drawing.Size(479, 17);
            this.cbSParseConsider.TabIndex = 7;
            this.cbSParseConsider.Text = "選択した一覧に /con, /whogroup, /whoraid コマンドでマークしたキャラクタを追加します。";
            this.cbSParseConsider.MouseHover += new System.EventHandler(this.cbSParseConsider_MouseHover);
            #if DEBUG_FILE_OUTPUT
            // 
            // cbDebugLog
            // 
            this.cbDebugLog.AutoSize = true;
            this.cbDebugLog.Checked = false;
            this.cbDebugLog.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbDebugLog.Location = new System.Drawing.Point(13, 109);
            this.cbDebugLog.Margin = new System.Windows.Forms.Padding(3, 1, 3, 1);
            this.cbDebugLog.Name = "cbDebugLog";
            this.cbDebugLog.Size = new System.Drawing.Size(677, 16);
            this.cbDebugLog.TabIndex = 10;
            this.cbDebugLog.Text = "（※デバッグ用機能なのでoff推奨。動作がかなり重くなります）JPNプラグインで解析できなかったログを指定ファイルに出力します。";
            this.cbDebugLog.MouseHover += new System.EventHandler(this.cbDebugLog_MouseHover);
            // 
            // tbDebugFileName
            // 
            this.tbDebugFileName.Location = new System.Drawing.Point(33, 128);
            this.tbDebugFileName.Name = "tbDebugFileName";
            this.tbDebugFileName.Size = new System.Drawing.Size(594, 19);
            this.tbDebugFileName.Text = @"D:\ACTJpnParser_debug.txt";
            this.tbDebugFileName.TabIndex = 11;
            // 
            // btDebugFileName
            // 
            this.btDebugFileName.Location = new System.Drawing.Point(626, 128);
            this.btDebugFileName.Name = "btDebugFileName";
            this.btDebugFileName.Size = new System.Drawing.Size(59, 21);
            this.btDebugFileName.TabIndex = 12;
            this.btDebugFileName.Text = "...";
            this.btDebugFileName.UseVisualStyleBackColor = true;
            #endif
            // 
            // cbIncludeInterceptFocus
            // 
            this.cbIncludeInterceptFocus.AutoSize = true;
            this.cbIncludeInterceptFocus.Location = new System.Drawing.Point(13, 90);
            this.cbIncludeInterceptFocus.Margin = new System.Windows.Forms.Padding(3, 1, 3, 1);
            this.cbIncludeInterceptFocus.Name = "cbIncludeInterceptFocus";
            this.cbIncludeInterceptFocus.Size = new System.Drawing.Size(466, 17);
            this.cbIncludeInterceptFocus.TabIndex = 19;
            this.cbIncludeInterceptFocus.Text = "チャネラーペットのフォーカスダメージを解析します。(攻撃者のDPSや命中率が正しく表示されなくなります。既に読み込んだデータには反映しません)";
            this.cbIncludeInterceptFocus.MouseHover += new System.EventHandler(this.cbIncludeInterceptFocus_MouseHover);
            // 
            // ACT_Jpn_Parser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            #if DEBUG_FILE_OUTPUT
            this.Controls.Add(this.btDebugFileName);
            this.Controls.Add(this.tbDebugFileName);
            this.Controls.Add(this.cbDebugLog);
            #endif
            this.Controls.Add(this.cbIncludeInterceptFocus);
            this.Controls.Add(this.cbKatakana);
            this.Controls.Add(this.cbMultiDamageIsOne);
            this.Controls.Add(this.cbSParseConsider);
            this.Controls.Add(this.cbRecalcWardedHits);
            this.Name = "ACT_Jpn_Parser";
            this.Size = new System.Drawing.Size(694, 108);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        #endregion
        public ACT_Jpn_Parser()
        {
            InitializeComponent();
        }

        Label lblStatus;    // The status label that appears in ACT's Plugin tab
        string settingsFile = Path.Combine(ActGlobals.oFormActMain.AppDataFolder.FullName, "Config\\ACT_Jpn_Parser.config.xml");
        private CheckBox cbMultiDamageIsOne;
        private CheckBox cbRecalcWardedHits;
        private CheckBox cbKatakana;
        SettingsSerializer xmlSettings;
        private CheckBox cbSParseConsider;
        private CheckBox cbIncludeInterceptFocus;
        #if DEBUG_FILE_OUTPUT
        private CheckBox cbDebugLog;
        private TextBox tbDebugFileName;
        private Button btDebugFileName;
        #endif
        TreeNode optionsNode = null;

        public void InitPlugin(TabPage pluginScreenSpace, Label pluginStatusText)
        {
            lblStatus = pluginStatusText;    // Hand the status label's reference to our local var
            pluginScreenSpace.Controls.Add(this);
            this.Dock = DockStyle.Fill;

            int dcIndex = -1;   // Find the Data Correction node in the Options tab
            for (int i = 0; i < ActGlobals.oFormActMain.OptionsTreeView.Nodes.Count; i++)
            {
                if (ActGlobals.oFormActMain.OptionsTreeView.Nodes[i].Text == "Data Correction")
                    dcIndex = i;
            }
            if (dcIndex != -1)
            {
                // Add our own node to the Data Correction node
                optionsNode = ActGlobals.oFormActMain.OptionsTreeView.Nodes[dcIndex].Nodes.Add("EQ2 Japanese Settings");
                // Register our user control(this) to our newly create node path.  All controls added to the list will be laid out left to right, top to bottom
                ActGlobals.oFormActMain.OptionsControlSets.Add(@"Data Correction\EQ2 Japanese Settings", new List<Control> { this });
                Label lblConfig = new Label();
                lblConfig.AutoSize = true;
                lblConfig.Text = "Option タブの Data Correction セクションの EQ2 Japanese Settings にてオプション設定ができます。";
                pluginScreenSpace.Controls.Add(lblConfig);
            }

            xmlSettings = new SettingsSerializer(this);    // Create a new settings serializer and pass it this instance
            LoadSettings();

            PopulateRegexArray();

            SetupEQ2EnglishEnvironment();

            ActGlobals.oFormActMain.BeforeLogLineRead += new LogLineEventDelegate(oFormActMain_BeforeLogLineRead);
            ActGlobals.oFormActMain.BeforeCombatAction += new CombatActionDelegate(oFormActMain_BeforeCombatAction);
            ActGlobals.oFormActMain.AfterCombatAction += new CombatActionDelegate(oFormActMain_AfterCombatAction);
            ActGlobals.oFormActMain.OnLogLineRead += new LogLineEventDelegate(oFormActMain_OnLogLineRead);
            //ActGlobals.oFormActMain.UpdateCheckClicked += new FormActMain.NullDelegate(oFormActMain_UpdateCheckClicked);
            //if (ActGlobals.oFormActMain.GetAutomaticUpdatesAllowed())   // If ACT is set to automatically check for updates, check for updates to the plugin
            //    new Thread(new ThreadStart(oFormActMain_UpdateCheckClicked)).Start();    // If we don't put this on a separate thread, web latency will delay the plugin init phase
            lblStatus.Text = "Plugin は有効です。";
        }

        public void DeInitPlugin()
        {
            ActGlobals.oFormActMain.BeforeLogLineRead -= oFormActMain_BeforeLogLineRead;
            ActGlobals.oFormActMain.BeforeCombatAction -= oFormActMain_BeforeCombatAction;
            ActGlobals.oFormActMain.AfterCombatAction -= oFormActMain_AfterCombatAction;
            ActGlobals.oFormActMain.OnLogLineRead -= oFormActMain_OnLogLineRead;
            //ActGlobals.oFormActMain.UpdateCheckClicked -= oFormActMain_UpdateCheckClicked;

            if (optionsNode != null)    // If we added our user control to the Options tab, remove it
            {
                optionsNode.Remove();
                ActGlobals.oFormActMain.OptionsControlSets.Remove(@"Data Correction\EQ2 Japanese Settings");
            }

            SaveSettings();
            lblStatus.Text = "Plugin は無効です。";    
        }

        #region oFormActMain_UpdateCheckClicked
        /*
        void oFormActMain_UpdateCheckClicked()
        {
            int pluginId = 55;
            try
            {
                DateTime localDate = ActGlobals.oFormActMain.PluginGetSelfDateUtc(this);
                DateTime remoteDate = ActGlobals.oFormActMain.PluginGetRemoteDateUtc(pluginId);
                if (localDate.AddHours(2) < remoteDate)
                {
                    DialogResult result = MessageBox.Show("There is an updated version of the EQ2 Japanese Parsing Plugin.  Update it now?\n\n(If there is an update to ACT, you should click No and update ACT first.)", "New Version", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        FileInfo updatedFile = ActGlobals.oFormActMain.PluginDownload(pluginId);
                        ActPluginData pluginData = ActGlobals.oFormActMain.PluginGetSelfData(this);
                        pluginData.pluginFile.Delete();
                        updatedFile.MoveTo(pluginData.pluginFile.FullName);
                        ThreadInvokes.CheckboxSetChecked(ActGlobals.oFormActMain, pluginData.cbEnabled, false);
                        Application.DoEvents();
                        ThreadInvokes.CheckboxSetChecked(ActGlobals.oFormActMain, pluginData.cbEnabled, true);
                    }
                }
            }
            catch (Exception ex)
            {
                ActGlobals.oFormActMain.WriteExceptionLog(ex, "Plugin Update Check");
            }
        }
        */
        #endregion

        #region Parsing
        Regex[] regexArray;
        const string logTimeStampRegexStr = @"\(\d{10}\)\[.{24}\] ";
        DateTime lastWardTime = DateTime.MinValue;
        int lastWardAmount = 0;
        string lastWardedTarget = string.Empty;
        DateTime lastInterceptTime = DateTime.MinValue;
        int lastInterceptAmount = 0;
        string lastInterceptTarget = string.Empty;
        string lastIntercepter = string.Empty;
        Regex petSplit = new Regex(@"(?<attacker>\w+)(?:'s|の) ?(?<petName>.+)", RegexOptions.Compiled);
        Regex romanjiSplit = new Regex(@"\\r[iapsmq]:(?<katakana>.+?)\\::?(?<romanji>.+)\\/r", RegexOptions.Compiled);
        Regex regexConsider = new Regex(logTimeStampRegexStr + @"(?:\\#......)?(?<player>.+?)を調べた。.+", RegexOptions.Compiled);
        Regex regexWhogroup = new Regex(logTimeStampRegexStr + @"(?<name>[^ ]+) Lvl \d+ .+", RegexOptions.Compiled);
        Regex regexWhoraid = new Regex(logTimeStampRegexStr + @"\[\d+ [^\]]+\] (?<name>[^ ]+) \([^\)]+\)", RegexOptions.Compiled);

        CombatActionEventArgs lastDamage = null;

        // -- Special Thanks Obanburumai  Start --
        Regex regexLogTimeStamp = new Regex(logTimeStampRegexStr, RegexOptions.Compiled);
        Regex regexSkillEol = new Regex(logTimeStampRegexStr + @"..*\\/r ?(?=\r?\n|$)", RegexOptions.Compiled);
        string skillEolLogline = string.Empty;
        bool isSkillEol = false;
        // -- Special Thanks Obanburumai  End --

        private void PopulateRegexArray()
        {
            regexArray = new Regex[23];
            regexArray[0]  = new Regex(logTimeStampRegexStr + @"(?<victim>.+?)(?:は|が) ?(?:(?<skillType>.+?)による攻撃を受け、)?(?<damageAndType>\d+(?:ポイント)?の ?)ダメージを負(?:いまし|っ)た。(?:[（(](?<special>.+?)[（)])?", RegexOptions.Compiled);
            regexArray[1]  = new Regex(logTimeStampRegexStr + @"(?<attacker>あなた|.+?)(?:は|が) ?(?<victim>.+?)(?:に ?|をヒット。)(?<damageAndType>.+?)の? ?ダメージ(?:を負わせ(?:まし)?た)?。(?:[（(](?<special>.+?)[）)])?", RegexOptions.Compiled);
            regexArray[2]  = new Regex(logTimeStampRegexStr + @"(?<attacker>あなた|.+?)(?:は|が) ?(?<victim>.+?)(?:に ?|をヒット。)(?<damageAndType>.+?)ポイント与え(?:まし)?た。(?:[（(](?<special>.+?)[）)])?", RegexOptions.Compiled);
            regexArray[3]  = new Regex(logTimeStampRegexStr + @"(?<victim>.+?)(?:が|は) ?(?<attacker>.+?)(?:の放った|の|'s) ?(?<skill>.+?)により(?<damageAndType>\d+(?:ポイントの| +).+?)ダメージを受け(?:た|ました)。(?:[（(](?<special>.+)[）)])?", RegexOptions.Compiled);
            regexArray[4]  = new Regex(logTimeStampRegexStr + @"(?<attacker>.+?)(?:の放った|の|'s) ?(?<skill>.+?)により、(?<victim>.+?)(?:が|は) ?(?<damageAndType>\d+(?:ポイントの| +).+?)ダメージを受け(?:た|ました)。(?:[（(](?<special>.+)[）)])?", RegexOptions.Compiled);
            regexArray[5]  = new Regex(logTimeStampRegexStr + @"(?<attacker>.+?)(?:の|'s) ?(?<skill>.+?)(?:で|があなたに命中し、)? ?(?<damageAndType>\d+ポイントの.+?)ダメージを受けましｈ?た。(?:[（(](?<special>.+?)[）)])?", RegexOptions.Compiled);
            regexArray[6]  = new Regex(logTimeStampRegexStr + @"(?<healer>.+?) ?(?:の|は|'s ?) ?(?<skill>.+)(?:が ?|によって、)(?<victim>.+?)を ?(?:(?:ヒール|修復|リペア)してい(?:ます|る)|回復させました)(?<damage>\d+) ?ヒットポイントの?(?<crit>(?:フェイブルド|レジェンダリ|ミシカル)?クリティカル)?。", RegexOptions.Compiled);
            regexArray[7]  = new Regex(logTimeStampRegexStr + @"(?<attacker>.+?)(?:が|は) ?(?<victim>.+?)を ?(?<skill>.+?で攻撃|攻撃|ヒット)(?:。はずした|(?:しましたが|しようとしましたが)、失敗しました)。(?:[（(](?<special>.+?)[）)])?", RegexOptions.Compiled);
            regexArray[8]  = new Regex(logTimeStampRegexStr + @"(?<attacker>.+?)(?:が|は) ?(?<victim>.+?)を ?攻撃(?:。|しましたが、)(?<why>.+?)(?:が ?うまく妨害|によって妨げられま|は ?うまくかわしま)した。(?:[（(].*(?<special>ブロック|反撃|回避|受け流し|レジスト|反射|強打|カウンター).*[）)])?", RegexOptions.Compiled);
            regexArray[9]  = new Regex(logTimeStampRegexStr + @"(?<attacker>.+?)(?:が|は) ?(?<victim>.+?)を ?(?<skill>.+?)で ?攻撃(?:しましたが、|。)(?<why>.+?)(?:が ?うまく妨害|によって妨げられま|は ?うまくかわしま)した。(?:[（(].*(?<special>ブロック|反撃|回避|受け流し|レジスト|反射|強打|カウンター).*[）)])?", RegexOptions.Compiled);
            regexArray[10] = new Regex(logTimeStampRegexStr + @"(?<victim>.+?)を ?倒した。", RegexOptions.Compiled);
            regexArray[11] = new Regex(logTimeStampRegexStr + @"(?<victim>.+?)?が ?(?<attacker>.+?)(?:(?:の|'s).+?)?に殺された……。", RegexOptions.Compiled);
            regexArray[12] = new Regex(logTimeStampRegexStr + @"(?<attacker>.+?)(?:(?:の|'s).+?)?に殺された……。", RegexOptions.Compiled);
            //regexArray[13] = new Regex(logTimeStampRegexStr + @"Unknown command: 'act (.+)'", RegexOptions.Compiled);
            regexArray[13] = new Regex(logTimeStampRegexStr + @"不明コマンド： 'act (.+)'", RegexOptions.Compiled);
            regexArray[14] = new Regex(logTimeStampRegexStr + @"(?<victim>.+?)は ?(?<attacker>[^\\].+?)の(?<skill>.+?) ?で攻撃を受け、ポイントパワーを(?<damage>\d+)(?:ポイント)?消耗し(?:た|ました)。(?:[（(](?<special>.+?)[）)])?", RegexOptions.Compiled);
            regexArray[15] = new Regex(logTimeStampRegexStr + @"(?:(?<attacker>[^\\].+?)(?:が|の)(?:放った)? ?)?(?:(?<skill>.+?) ?(?:で ?|が ?|により))?(?<victim>.+?)?(?:を ?攻撃し|に ?命中し|に ?ヒットし|攻撃を受け)、(?:ポイントパワーを)?(?<damage>\d+)ポイント(?:のポイント)?(?:パワーを)?消耗(?:させ|し)(?:た|ました)。?(?:[（(](?<special>.+?)[）)])?", RegexOptions.Compiled);
            regexArray[16] = new Regex(logTimeStampRegexStr + @"(?<victim>.+)に対する(?<damage>\d+) ?ポイントダメージを(?<attacker>あなた|.+?)(?:の|'s) ?(?<skillType>.+?)が吸収した。", RegexOptions.Compiled);
            regexArray[17] = new Regex(logTimeStampRegexStr + @"(?<skill>.+)は ?(?<damage>\d+) ?ポイントのダメージを吸収し、(?<victim>.+?)へのダメージを防いだ(?:。)?", RegexOptions.Compiled);
            regexArray[18] = new Regex(logTimeStampRegexStr + @"You have entered (?<zone>.+?)\.", RegexOptions.Compiled);
            regexArray[19] = new Regex(logTimeStampRegexStr + @"(?<healer>.+?)(?:が|は|の|'s) ?(?<skill>.+?)(?:が|で|によって) ?(?<victim>.+?)をリフレッシュしてい(?:る|ます)(?<damage>\d+) ?マナポイント(?:の)?(?<special>(?:フェイブルド|レジェンダリ|ミシカル)?クリティカル)?。", RegexOptions.Compiled);
            regexArray[20] = new Regex(logTimeStampRegexStr + @"(?:(?<owner>[^\\].+?|あなた)(?:の|'s))? ?(?<skillType>.+?)(?:が|で) ?(?<victim>.+?)に対する(?<target>.*)のヘイト(?:順位)?を ?(?<direction>増加|減少) ?(?<damage>\d+) ?(?<dirType>脅威レベル|position)の?(?<crit>(?:フェイブルド|レジェンダリ|ミシカル)?クリティカル)?。", RegexOptions.Compiled);
            regexArray[21] = new Regex(logTimeStampRegexStr + @"(?<attacker>.+?|あなた)(?:の|'s) ?(?<skillType>.+?)が ?(?:(?<victim>.+?|あなた))の ?(?<affliction>.+?)を ?(?<action>ディスペル|治療)しました。", RegexOptions.Compiled);
            regexArray[22] = new Regex(logTimeStampRegexStr + @"(?<healer>.+?)[はが] ?(?<attacker>.+?)から(?<victim>.+?)へのダメージを ?(?<damage>\d+) ?減らし(?:まし)?た。", RegexOptions.Compiled);
        }
        void oFormActMain_BeforeLogLineRead(bool isImport, LogLineEventArgs logInfo)
        {
            if (NotQuickFail(logInfo))
            {
                #if DEBUG_FILE_OUTPUT
                bool isMatchLog = false;
                #endif

                // -- Special Thanks Obanburumai  Start --
                if (this.isSkillEol)
                {
                    if (!regexLogTimeStamp.IsMatch(logInfo.logLine))
                        logInfo.logLine = this.skillEolLogline + logInfo.logLine;

                    this.skillEolLogline = string.Empty;
                    this.isSkillEol = false;
                }
                // -- Special Thanks Obanburumai  End --

                for (int i = 0; i < regexArray.Length; i++)
                {
                    Match reMatch = regexArray[i].Match(logInfo.logLine);
                    if (reMatch.Success)
                    {
                        #if DEBUG_FILE_OUTPUT
                        isMatchLog = true;
                        #endif
                        switch (i)
                        {
                            case 0:
                            case 1:
                            case 2:
                            case 3:
                            case 4:
                            case 5:
                                logInfo.detectedType = Color.Red.ToArgb();
                                break;
                            case 6:
                                logInfo.detectedType = Color.Blue.ToArgb();
                                break;
                            case 7:
                            case 8:
                            case 9:
                                logInfo.detectedType = Color.DarkRed.ToArgb();
                                break;
                            case 14:
                            case 15:
                                logInfo.detectedType = Color.DarkOrchid.ToArgb();
                                break;
                            case 16:
                            case 17:
                                logInfo.detectedType = Color.DodgerBlue.ToArgb();
                                break;
                            default:
                                logInfo.detectedType = Color.Black.ToArgb();
                                break;
                        }
                        LogExeJpn(reMatch, i + 1, logInfo.logLine, isImport);
                        break;
                    }
                }
                #if DEBUG_FILE_OUTPUT
                if (!isMatchLog && this.cbDebugLog.Checked)
                {
                    string filename = @"D:\ACTJpnParser_debug.txt";
                    if (this.tbDebugFileName.Text.Length > 0) filename = this.tbDebugFileName.Text;
                    StreamWriter writer = new StreamWriter(filename, true, Encoding.UTF8);
                    writer.WriteLine(logInfo.logLine);
                    writer.Close();
                }
                #endif
            }

            // -- Special Thanks Obanburumai  Start --
            if (regexSkillEol.IsMatch(logInfo.logLine))
            {
                this.isSkillEol = true;
                this.skillEolLogline = logInfo.logLine;
                this.skillEolLogline = this.skillEolLogline.Replace("\r", "").Replace("\n", "");
            }
            // -- Special Thanks Obanburumai  End --
        }
        void oFormActMain_BeforeCombatAction(bool isImport, CombatActionEventArgs actionInfo)
        {
            // Riposte/kontert/отвечает & Relect/отражает
            if (lastDamage != null && lastDamage.time == actionInfo.time)
            {
                if ((int)lastDamage.damage == (int)Dnum.Unknown && lastDamage.damage.DamageString.Contains("反撃"))
                {
                    if (actionInfo.swingType == (int)SwingTypeEnum.Melee && actionInfo.victim == lastDamage.attacker)
                    {
                        actionInfo.special = "反撃";
                        lastDamage.damage.DamageString2 = String.Format("({0} returned)", actionInfo.damage.ToString());
                    }
                }
                if ((int)actionInfo.damage == (int)Dnum.Unknown && actionInfo.damage.DamageString.Contains("反射"))
                {
                    if (actionInfo.theAttackType == lastDamage.theAttackType && actionInfo.victim == lastDamage.attacker)
                    {
                        //lastDamage.special = "reflect";  // Too late to take effect
                        actionInfo.damage.DamageString2 = String.Format(" ({0} returned)", lastDamage.damage.ToString());
                    }
                }
            }
        }
        void oFormActMain_AfterCombatAction(bool isImport, CombatActionEventArgs actionInfo)
        {
            if (actionInfo.swingType == (int)SwingTypeEnum.Melee || actionInfo.swingType == (int)SwingTypeEnum.NonMelee)
                lastDamage = actionInfo;
        }
        bool captureWhoraid = false;
        void oFormActMain_OnLogLineRead(bool isImport, LogLineEventArgs logInfo)
        {
            if (cbSParseConsider.Checked)
            {
                if (logInfo.logLine.EndsWith("に対する/whoraidの結果"))
                    captureWhoraid = true;
                if (logInfo.logLine.EndsWith("人のプレイヤーがいます。"))
                    captureWhoraid = false;
                if (captureWhoraid && regexWhoraid.IsMatch(logInfo.logLine))
                {
                    string outputName = regexWhoraid.Replace(logInfo.logLine, "$1");
                    ActGlobals.oFormActMain.SelectiveListAdd(outputName);
                }
                if (regexWhogroup.IsMatch(logInfo.logLine))
                {
                    string outputName = regexWhogroup.Replace(logInfo.logLine, "$1");
                    ActGlobals.oFormActMain.SelectiveListAdd(outputName);
                }
                if (regexConsider.IsMatch(logInfo.logLine))
                {
                    string outputName = regexConsider.Replace(logInfo.logLine, "$1");
                    ActGlobals.oFormActMain.SelectiveListAdd(outputName);
                    if (!isImport)
                        System.Media.SystemSounds.Beep.Play();
                }
            }
        }

        string[] matchKeywords = new string[] { "ダメージ", "ヒットポイント", "マナポイント", "はずした", "失敗しました", "妨害した", "妨げられました", "かわしました", "倒した", "殺された", "不明コマンド： 'act", "entered", "消耗", "ヘイト", "ディスペル", "治療" };
        private bool NotQuickFail(LogLineEventArgs logInfo)
        {
            foreach (string s in matchKeywords)
            {
                if (logInfo.logLine.Contains(s))
                    return true;
            }

            return false;
        }

        private void LogExeJpn(Match reMatch, int logMatched, string logLine, bool isImport)
        {
            string attacker, victim, damage, skillType, why, crit, special, attackType;
            List<string> attackingTypes = new List<string>();
            List<string> damages = new List<string>();
            Regex rE = regexArray[logMatched - 1];
            SwingTypeEnum swingType;
            bool critical = false;
            List<DamageAndType> damageAndTypeArr;

            DateTime time = ActGlobals.oFormActMain.LastKnownTime;

            Dnum failType;
            int gts = ActGlobals.oFormActMain.GlobalTimeSorter;

            switch (logMatched)
            {
                #region Case 1 [unsourced skill attacks]
                case 1:
                    victim = reMatch.Groups[1].Value;
                    skillType = reMatch.Groups[2].Value;
                    skillType = String.IsNullOrEmpty(skillType) ? "不明" : skillType;
                    damage = reMatch.Groups[3].Value;
                    special = reMatch.Groups[4].Value;
                    special = String.IsNullOrEmpty(special) ? "None" : special;
                    attacker = "不明";    // Unsourced melee hits show as "Unknown" attacking, so we do the same
                    if (!ActGlobals.oFormActMain.InCombat && !isImport)
                    {
                        ActGlobals.oFormSpellTimers.NotifySpell(attacker.ToLower(), skillType, victim.Contains("あなた"), victim.ToLower(), true);
                        break;
                    }
                    damageAndTypeArr = JpnGetDamageAndTypeArr(damage);

                    if (victim == "あなた")
                        victim = ActGlobals.charName;
                    AddDamageAttack(attacker, victim, skillType, (int)SwingTypeEnum.NonMelee, critical, special, damageAndTypeArr, time, gts);
                    break;
                #endregion
                #region Case 2 [melee/non-melee attacks]
                case 2:
                case 3:
                case 4:
                case 5:
                case 6:
                    if (logMatched == 2) {
                        attacker = reMatch.Groups[1].Value;
                        victim = reMatch.Groups[2].Value;
                        damage = reMatch.Groups[3].Value;
                        special = reMatch.Groups[4].Value;
                        crit = special;
                        if (petSplit.IsMatch(attacker)) {
                            skillType = petSplit.Replace(attacker, "$2") + "の攻撃";
                            attacker = petSplit.Replace(attacker, "$1");
                        }
                        else skillType = "攻撃";
                    }
                    else if (logMatched == 3) {
                        attacker = reMatch.Groups[1].Value;
                        victim = reMatch.Groups[2].Value;
                        damage = reMatch.Groups[3].Value;
                        special = reMatch.Groups[4].Value;
                        crit = special;
                        if (petSplit.IsMatch(attacker)) {
                            skillType = petSplit.Replace(attacker, "$2") + "の攻撃";
                            attacker = petSplit.Replace(attacker, "$1");
                        }
                        else skillType = "攻撃";
                    }
                    else if (logMatched == 4) {
                        victim = reMatch.Groups[1].Value;
                        attacker = reMatch.Groups[2].Value;
                        skillType = reMatch.Groups[3].Value;
                        damage = reMatch.Groups[4].Value;
                        special = reMatch.Groups[5].Value;
                        crit = special;
                    }
                    else if (logMatched == 5) {
                        attacker = reMatch.Groups[1].Value;
                        skillType = reMatch.Groups[2].Value;
                        victim = reMatch.Groups[3].Value;
                        damage = reMatch.Groups[4].Value;
                        special = reMatch.Groups[5].Value;
                        crit = special;
                    }
                    else {
                        attacker = reMatch.Groups[1].Value;
                        skillType = reMatch.Groups[2].Value;
                        damage = reMatch.Groups[3].Value;
                        special = reMatch.Groups[4].Value;
                        victim = "あなた";
                        crit = special;
                    }

                    if (damage.Contains("pain and suffering"))
                        break;
                    damageAndTypeArr = JpnGetDamageAndTypeArr(damage);

                    special = String.IsNullOrEmpty(special) ? "None" : special;

                    victim = JapanesePersonaReplace(victim);
                    if (attacker == "あなた")    // You performing melee
                        attacker = ActGlobals.charName;

                    if (attacker == victim || attacker == petSplit.Replace(victim, "$2"))
                        break;        // You don't get credit for attacking yourself or your own pet
                    // Traumatic Swipe / トラウマティック・スワイプ
                    if (skillType == "Traumatic Swipe" || skillType == "トラウマティック・スワイプ" || skillType == "\\ra:トラウマティック・スワイプ\\:Traumatic Swipe\\/r")
                        ActGlobals.oFormSpellTimers.ApplyTimerMod(attacker, victim, skillType, 0.5F, 30);
                    if (skillType == "攻撃") {
                        swingType = SwingTypeEnum.Melee;
                        skillType = "";
                    }
                    else
                        swingType = SwingTypeEnum.NonMelee;

                    if (ActGlobals.oFormActMain.SetEncounter(time, attacker, victim))
                    {
                        AddDamageAttack(attacker, victim, skillType, (int)swingType, critical, special, damageAndTypeArr, time, gts);
                    }
                    //NotifySpell(attacker, skillType, true, victim, true);

                    break;
                #endregion
                #region Case 3 [healing]
                case 7:
                    if (!ActGlobals.oFormActMain.InCombat)
                        break;
                    attacker = reMatch.Groups[1].Value;
                    skillType = reMatch.Groups[2].Value;
                    victim = reMatch.Groups[3].Value;
                    damage = reMatch.Groups[4].Value;
                    crit = reMatch.Groups[5].Value;

                    if (crit.Contains("クリティカル"))        // Check for critical hits
                        critical = true;
                    else
                        critical = false;

                    special = crit.Replace("クリティカル", string.Empty).Trim();
                    special = String.IsNullOrEmpty(special) ? "None" : special;

                    victim = JapanesePersonaReplace(victim);
                    if (attacker == "あなた")        // You healing
                        attacker = ActGlobals.charName;

                    AddCombatActionTrans((int)SwingTypeEnum.Healing, critical, special, attacker, skillType, new Dnum(Int32.Parse(damage)), time, gts, victim, "Hitpoints");
                    break;
                #endregion
                #region Case 4 [misses]
                case 8:
                case 9:
                case 10:
                    string damageType = "";
                    if (logMatched == 8) {
                        attacker = reMatch.Groups[1].Value;
                        victim = reMatch.Groups[2].Value;
                        why = reMatch.Groups[3].Value;
                        special = reMatch.Groups[4].Value;
                        if (((why == "攻撃" || why == "ヒット")) && (!petSplit.IsMatch(attacker)))
                        {
                            skillType = "攻撃";
                            why = "";
                        }
                        else { // skill miss
                            skillType = (Regex.Split( why , "で攻撃" ))[0].Trim();
                            if (petSplit.IsMatch(attacker))
                            {
                                skillType = petSplit.Replace(attacker, "$2") + "の" + skillType;
                                attacker = petSplit.Replace(attacker, "$1");
                            }
                        }
                    }
                    else if (logMatched == 9) {
                        attacker = reMatch.Groups[1].Value;
                        victim = reMatch.Groups[2].Value;
                        why = reMatch.Groups[3].Value;
                        special = reMatch.Groups[4].Value;
                        skillType = "攻撃";
                    }
                    else {
                        attacker = reMatch.Groups[1].Value;
                        victim = reMatch.Groups[2].Value;
                        skillType = reMatch.Groups[3].Value;
                        why = reMatch.Groups[4].Value;
                        special = reMatch.Groups[5].Value;
                    }

                    attacker = JapanesePersonaReplace(attacker);
                    victim = JapanesePersonaReplace(victim);

                    if (skillType == "攻撃") {
                        damageType = "melee";
                        swingType = SwingTypeEnum.Melee;
                    }
                    else {
                        damageType = "non-melee";
                        swingType = SwingTypeEnum.NonMelee;
                    }

                    why = why.Replace(victim, string.Empty);

                    if (String.IsNullOrEmpty(why))
                        failType = Dnum.Miss;
                    else
                    {
                        why = why.Trim() + " " + special;
                        failType = new Dnum(Dnum.Unknown, why.Trim());
                    }

                    special = String.IsNullOrEmpty(special) ? "None" : special;

                    if (attacker == victim || attacker == petSplit.Replace(victim, "$2"))
                        break;        // You don't get credit for attacking yourself or your own pet

                    if (ActGlobals.oFormActMain.SetEncounter(time, attacker, victim))
                        AddCombatActionTrans((int)swingType, false, special, attacker, skillType, failType, time, gts, victim, damageType);
                    break;
                #endregion
                #region Case 5 [killing]
                case 11:
                case 12:
                case 13:
                    if (logMatched == 11) {
                        attacker = "あなた";
                        victim = reMatch.Groups[1].Value;
                    }
                    else if (logMatched == 12) {
                        victim = reMatch.Groups[1].Value;
                        attacker = reMatch.Groups[2].Value;
                    }
                    else {
                        victim = "あなた";
                        attacker = reMatch.Groups[1].Value;
                    }

                    if (victim == "あなた")
                        victim = ActGlobals.charName;
                    else if (victim == String.Empty)
                        victim = "不明";
                    if (attacker == "あなた")
                        attacker = ActGlobals.charName;
                    else if (attacker == String.Empty)
                        attacker = "不明";

                    swingType = SwingTypeEnum.NonMelee;
                    ActGlobals.oFormSpellTimers.RemoveTimerMods(victim);
                    ActGlobals.oFormSpellTimers.DispellTimerMods(victim);
                    if (ActGlobals.oFormActMain.InCombat)
                    {
                        AddCombatActionTrans((int)swingType, false, "None", attacker, "Killing", Dnum.Death, time, gts, victim, "Death");
                    }
                    break;
                #endregion
                #region Case 6 [act commands]
                case 14:
                    ActGlobals.oFormActMain.ActCommands(rE.Replace(logLine, "$1"));
                    break;
                #endregion
                #region Case 7 [power drain]
                case 15:
                case 16:
                    if (logMatched == 15) {
                        victim = reMatch.Groups[1].Value;
                        attacker = reMatch.Groups[2].Value;
                        skillType = reMatch.Groups[3].Value;
                        damage = reMatch.Groups[4].Value;
                        crit = reMatch.Groups[5].Value;
                    }
                    else {
                        attacker = reMatch.Groups[1].Value;
                        skillType = reMatch.Groups[2].Value;
                        victim = reMatch.Groups[3].Value;
                        damage = reMatch.Groups[4].Value;
                        crit = reMatch.Groups[5].Value;
                        if (attacker == String.Empty)
                            attacker = "あなた";
                        if (victim == String.Empty)
                            victim = "あなた";
                    }

                    attacker = JapanesePersonaReplace(attacker);
                    victim = JapanesePersonaReplace(victim);
                    critical = crit.Contains("クリティカル");

                    if (attacker == victim || attacker == petSplit.Replace(victim, "$2"))
                        break;        // You don't get credit for attacking yourself or your own pet
                    if (ActGlobals.oFormActMain.SetEncounter(time, attacker, victim))
                    {
                        if (CheckWardedHit(victim, time))
                        {
                            Dnum complexWardedHit = new Dnum(Int32.Parse(damage) + lastWardAmount, String.Format("{0}/{1}", lastWardAmount, damage));
                            AddCombatActionTrans((int)SwingTypeEnum.PowerDrain, false, "None", attacker, skillType, complexWardedHit, time, gts, victim, "warded/non-melee");
                            lastWardAmount = 0;
                        }
                        else
                            AddCombatActionTrans((int)SwingTypeEnum.PowerDrain, false, "None", attacker, skillType, new Dnum(Int32.Parse(damage)), time, gts, victim, "non-melee");
                    }
                    break;
                #endregion
                #region Case 8 [ward absorbtion]
                case 17:
                case 18:
                    if (!ActGlobals.oFormActMain.InCombat)
                        break;
                    if (logMatched == 17) {
                        victim = reMatch.Groups[1].Value;
                        damage = reMatch.Groups[2].Value;
                        attacker = reMatch.Groups[3].Value;
                        skillType = reMatch.Groups[4].Value;
                    }
                    else {
                        skillType = reMatch.Groups[1].Value;
                        damage = reMatch.Groups[2].Value;
                        victim = reMatch.Groups[3].Value;
                        attacker = "あなた";
                    }

                    attacker = JapanesePersonaReplace(attacker);
                    victim = JapanesePersonaReplace(victim);

                    AddCombatActionTrans((int)SwingTypeEnum.Healing, false, "None", attacker, skillType, new Dnum(Int32.Parse(damage)), time, gts, victim, "Absorption");

                    if (CheckWardedHit(victim, time))
                        lastWardAmount += Int32.Parse(damage);
                    else
                        lastWardAmount = Int32.Parse(damage);
                    lastWardedTarget = victim;
                    lastWardTime = time;
                    break;
                #endregion
                #region Case 9 [zone change]
                case 19:
                    if (logLine.Contains(" combat by "))
                        break;
                    ActGlobals.oFormActMain.ChangeZone(rE.Replace(logLine, "$1").Trim());
                    break;
                #endregion
                #region Case 10 [power healing]
                case 20:
                    if (!ActGlobals.oFormActMain.InCombat)
                        break;
                    attacker = reMatch.Groups[1].Value;
                    skillType = reMatch.Groups[2].Value;
                    victim = reMatch.Groups[3].Value;
                    damage = reMatch.Groups[4].Value;
                    crit = reMatch.Groups[5].Value;

                    if (crit.Contains("クリティカル"))        // Check for critical hits
                        critical = true;
                    else
                        critical = false;

                    victim = JapanesePersonaReplace(victim);
                    if (attacker.StartsWith("あなた"))        // You healing
                        attacker = ActGlobals.charName;
                    AddCombatActionTrans((int)SwingTypeEnum.PowerHealing, critical, "None", attacker, skillType, new Dnum(Int32.Parse(damage)), time, gts, victim, "Power");
                    break;
                #endregion
                #region Case 11 [threat]
                case 21:
                    attacker = reMatch.Groups[1].Value;
                    skillType = reMatch.Groups[2].Value;
                    victim = reMatch.Groups[3].Value;
                    string owner = reMatch.Groups[4].Value;
                    string direction = reMatch.Groups[5].Value;
                    damage = reMatch.Groups[6].Value;
                    string position = reMatch.Groups[7].Value;
                    special = reMatch.Groups[8].Value;

                    attacker = JapanesePersonaReplace(attacker);
                    victim = JapanesePersonaReplace(victim);

                    if (String.IsNullOrEmpty(attacker))
                        attacker = ActGlobals.charName;
                    if (owner == "相手")
                        owner = attacker;

                    bool increase = (direction == "増加");
                    bool positionChange = (position == "position");
                    special = String.IsNullOrEmpty(special) ? "None" : special;

                    Dnum dDamage;
                    if (positionChange)
                        dDamage = new Dnum(Dnum.ThreatPosition, String.Format("{0} Positions", Int32.Parse(damage)));
                    else
                        dDamage = new Dnum(Int32.Parse(damage));
                    direction = increase ? "Increase" : "Decrease";

                    if (attacker == victim || attacker == petSplit.Replace(victim, "$2"))
                        break;        // You don't get credit for attacking yourself or your own pet
                    if (ActGlobals.oFormActMain.SetEncounter(time, attacker, victim) || ActGlobals.oFormActMain.SetEncounter(time, owner, victim))
                        AddCombatActionTrans((int)SwingTypeEnum.Threat, critical, special, attacker, skillType, dDamage, time, gts, victim, direction);
                    break;
                #endregion
                #region Case 12 [dispell/cure]
                case 22:
                    attacker = reMatch.Groups[1].Value;
                    skillType = reMatch.Groups[2].Value;
                    victim = reMatch.Groups[3].Value;
                    attackType = reMatch.Groups[4].Value;
                    direction = reMatch.Groups[5].Value;

                    attacker = JapanesePersonaReplace(attacker);
                    victim = JapanesePersonaReplace(victim);

                    if (attackType == "Traumatic Swipe" || attackType == "トラウマティック・スワイプ" || attackType == "\\ra:トラウマティック・スワイプ\\:Traumatic Swipe\\/r")
                        ActGlobals.oFormSpellTimers.DispellTimerMods(victim);

                    bool cont = false;
                    if (direction == "治療")
                        cont = ActGlobals.oFormActMain.InCombat;
                    else
                        cont = ActGlobals.oFormActMain.SetEncounter(time, attacker, victim);
                    if (cont)
                        AddCombatActionTrans((int)SwingTypeEnum.CureDispel, false, attackType, attacker, skillType, new Dnum(1), time, gts, victim, direction);
                    break;
                #endregion
                #region Case 13 [damage interception]
                case 23:
                    if (!ActGlobals.oFormActMain.InCombat)
                        break;
                    attacker = reMatch.Groups[1].Value;    // Inteceptor
                    special = reMatch.Groups[2].Value;    // Attacker
                    victim = reMatch.Groups[3].Value;    // Target
                    damage = reMatch.Groups[4].Value;    // Amount

                    attacker = JapanesePersonaReplace(attacker);
                    victim = JapanesePersonaReplace(victim);

                    AddCombatActionTrans((int)SwingTypeEnum.Healing, false, special, attacker, "Channeler Pet", new Dnum(Int32.Parse(damage)), time, gts, victim, "Interception");
                    if (CheckInterceptedHit(victim, time))
                        lastInterceptAmount += Int32.Parse(damage);
                    else
                        lastInterceptAmount = Int32.Parse(damage);
                    lastInterceptTarget = victim;
                    lastInterceptTime = time;
                    lastIntercepter = attacker;
                    break;
                #endregion
                default:
                    #if DEBUG_FILE_OUTPUT
                    string filename = @"D:\ACTJpnParser_debug.txt";
                    if (this.tbDebugFileName.Text.Length > 0) filename = this.tbDebugFileName.Text;
                    StreamWriter writer = new StreamWriter(filename, true, Encoding.UTF8);
                    writer.WriteLine(logLine);
                    writer.Close();
                    #endif
                    break;
            }
        }


        private bool CheckInterceptedFocus(string victim, DateTime time, List<DamageAndType> damageAndTypeArr)
        {
            if (cbRecalcWardedHits.Checked && lastInterceptTime == time && lastIntercepter == victim && lastInterceptAmount > 0)
            {
                if (damageAndTypeArr.Count == 1)
                {
                    if (damageAndTypeArr[0].Type == "フォーカス")
                        return true;
                }
            }
            return false;
        }
        private bool CheckInterceptedHit(string victim, DateTime time)
        {
            return cbRecalcWardedHits.Checked && lastInterceptTime == time && lastInterceptTarget == victim && lastInterceptAmount > 0;
        }
        private bool CheckWardedHit(string victim, DateTime time)
        {
            return cbRecalcWardedHits.Checked && lastWardTime == time && lastWardedTarget == victim && lastWardAmount > 0;
        }
        private void AddDamageAttack(string attacker, string victim, string skillType, int swingType, bool critical, string special, List<DamageAndType> damageAndTypeArr, DateTime time, int gts)
        {
            int damageTotal = 0;
            if (cbMultiDamageIsOne.Checked)
            {
                string damageStr = string.Empty;
                string typeStr = string.Empty;
                if (CheckInterceptedFocus(victim, time, damageAndTypeArr))
                {
                    if (!cbIncludeInterceptFocus.Checked)
                        return;
                }
                if (CheckInterceptedHit(victim, time))
                {
                    damageTotal = lastInterceptAmount;
                    damageStr += String.Format("{0}/", damageTotal);
                    typeStr += String.Format("{0}/", "intercepted");
                    lastInterceptAmount = 0;
                }
                if (CheckWardedHit(victim, time))
                {
                    damageTotal = lastWardAmount;
                    damageStr += String.Format("{0}/", damageTotal);
                    typeStr += String.Format("{0}/", "warded");
                    lastWardAmount = 0;
                }
                for (int i = 0; i < damageAndTypeArr.Count; i++)
                {
                    damageTotal += damageAndTypeArr[i].Damage;
                    damageStr += String.Format("{0}/", damageAndTypeArr[i].Damage);
                    typeStr += String.Format("{0}/", damageAndTypeArr[i].Type);
                }
                damageStr = damageStr.TrimEnd(new char[] { '/' });
                typeStr = typeStr.TrimEnd(new char[] { '/' });
                if (String.IsNullOrEmpty(skillType))
                    skillType = typeStr;

                AddCombatActionTrans(swingType, critical, special, attacker, skillType, new Dnum(damageTotal, damageStr), time, gts, victim, typeStr);
            }
            else
            {
                bool nullSkillType = String.IsNullOrEmpty(skillType);

                for (int i = 0; i < damageAndTypeArr.Count; i++)
                {
                    damageTotal = damageAndTypeArr[i].Damage;
                    string damageStr = damageAndTypeArr[i].Damage.ToString();
                    if (nullSkillType)
                        skillType = damageAndTypeArr[i].Type;

                    if (CheckInterceptedFocus(victim, time, damageAndTypeArr))
                    {
                        if (!cbIncludeInterceptFocus.Checked)
                            continue;
                    }

                    if (i == damageAndTypeArr.Count - 1 && (CheckWardedHit(victim, time) || CheckInterceptedHit(victim, time)))
                    {
                        damageTotal += lastInterceptAmount;
                        damageTotal += lastWardAmount;
                        lastInterceptAmount = 0;
                        lastWardAmount = 0;
                    }
                    AddCombatActionTrans(swingType, critical, special, attacker, skillType, new Dnum(damageTotal, damageStr), time, gts, victim, damageAndTypeArr[i].Type);
                }
            }
        }
        public void AddCombatAction(int SwingType, bool Critical, string Special, string Attacker, string theAttackType, Dnum Damage, DateTime Time, int TimeSorter, string Victim, string theDamageType)
        {
            if (romanjiSplit.IsMatch(theAttackType))
            {
                if (cbKatakana.Checked)
                    theAttackType = romanjiSplit.Replace(theAttackType, "$1");
                else
                    theAttackType = romanjiSplit.Replace(theAttackType, "$2");
            }
            if (romanjiSplit.IsMatch(Attacker))
            {
                if (cbKatakana.Checked)
                    Attacker = romanjiSplit.Replace(Attacker, "$1");
                else
                    Attacker = romanjiSplit.Replace(Attacker, "$2");
            }
            if (romanjiSplit.IsMatch(Victim))
            {
                if (cbKatakana.Checked)
                    Victim = romanjiSplit.Replace(Victim, "$1");
                else
                    Victim = romanjiSplit.Replace(Victim, "$2");
            }

            ActGlobals.oFormActMain.AddCombatAction(SwingType, Critical, Special, Attacker, theAttackType, Damage, Time, TimeSorter, Victim, theDamageType);
        }

        string[] criticalWords = new string[] { "レジェンダリ ", "フェイブルド ", "ミシカル " };
        public void AddCombatActionTrans(int SwingType, bool Critical, string Special, string Attacker, string theAttackType, Dnum Damage, DateTime Time, int TimeSorter, string Victim, string theDamageType)
        {
            string critStr = String.Empty;
            if (Special.Contains("クリティカル"))
            {
                Critical = true;
                foreach (string s in criticalWords)
                {
                    if (Special.Contains(s))
                    {
                        Special = Special.Replace(s, string.Empty);
                        critStr = s;
                        break;
                    }
                }
                Special = Special.Replace("・", string.Empty);
                Special = Special.Replace("ヒット", string.Empty);
                Special = Special.Replace("クリティカル", string.Empty);
                Special = Special.Replace("Critical ", string.Empty);
                critStr += "クリティカルヒット";
            }

            if (!Critical)
                critStr = "None";

            Special = String.IsNullOrEmpty(Special) ? "None" : Special;

            if (romanjiSplit.IsMatch(theDamageType))
                theDamageType = translateForMultiple(theDamageType);

            if (romanjiSplit.IsMatch(theAttackType))
                theAttackType = translateForMultiple(theAttackType);

            if (romanjiSplit.IsMatch(Attacker))
                Attacker = translateForMultiple(Attacker);

            if (romanjiSplit.IsMatch(Victim))
                Victim = translateForMultiple(Victim);

            if (romanjiSplit.IsMatch(Special))
                Special = translateForMultiple(Special);

            MasterSwing ms = new MasterSwing(SwingType, Critical, Special, Damage, Time, TimeSorter, theAttackType, Attacker, theDamageType, Victim);
            ms.Tags["CriticalStr"] = critStr;
            ActGlobals.oFormActMain.AddCombatAction(ms);
        }

        private string translateForMultiple(string transTarget)
        {
            string returnValue = string.Empty;

            transTarget = transTarget.Replace("\\ri:","");
            transTarget = transTarget.Replace("\\ra:","");
            transTarget = transTarget.Replace("\\rp:","");
            transTarget = transTarget.Replace("\\rm:","");
            transTarget = transTarget.Replace("\\rs:","");
            transTarget = transTarget.Replace("\\rq:","");
            transTarget = transTarget.Replace("\\::",":");
            transTarget = transTarget.Replace("\\:",":");
            transTarget = transTarget.Replace("\\/r",":");

            string[] arrTarget = transTarget.Split(':');

            StringBuilder jpn = new StringBuilder();
            StringBuilder eng = new StringBuilder();

            if(arrTarget.Length>1){
                for(int i=0;i<arrTarget.Length;i++){
                    if(i%2==0){
                        jpn.Append(arrTarget[i]);
                    }else{
                        eng.Append(arrTarget[i]);
                    }
                }

                if (cbKatakana.Checked){
                    returnValue = jpn.ToString();
                }else{
                    returnValue = eng.ToString();
                }
            }else{
                returnValue = transTarget;
            }

            return returnValue;
        }

        private string JapanesePersonaReplace(string input)
        {
            if (input.Contains("あなた"))
                return ActGlobals.charName;
            if (input.Contains("自分"))
                return ActGlobals.charName;
            return input;
        }

        Regex typeDamageSplit = new Regex(@"(?<type>.+)ダメージを(?<damage>\d+)", RegexOptions.Compiled);
        private List<DamageAndType> JpnGetDamageAndTypeArr(string damageAndType)
        {
            List<DamageAndType> outList = new List<DamageAndType>();
            damageAndType = damageAndType.Replace(" and ", ", ");
            damageAndType = damageAndType.Replace("ポイントの", " ");
            string[] entries = damageAndType.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < entries.Length; i++)
            {
                if (typeDamageSplit.IsMatch(entries[i]))
                	entries[i] = typeDamageSplit.Replace(entries[i], "$2") + " " + typeDamageSplit.Replace(entries[i], "$1");
                outList.Add(new DamageAndType(entries[i]));
            }
            return outList;
        }

        private class DamageAndType
        {
            int damage;
            string type;
            /// <summary>
            /// Data class for a single type of damage and the amount
            /// </summary>
            /// <param name="Damage">The positive integer amount of damage</param>
            /// <param name="Type">The type of damage to display it as</param>
            public DamageAndType(int Damage, string Type)
            {
                this.damage = Damage;
                this.type = Type;
            }
            /// <summary>
            /// Data class for a single type of damage and the amount
            /// </summary>
            /// <param name="UnsplitSource">An input string such as "123 crushing" to be split by the constructor</param>
            public DamageAndType(string UnsplitSource)
            {
                int spacePos = UnsplitSource.IndexOf(' ');
                if (spacePos == -1)
                    throw new ArgumentException("The input string did not contain a space, thus cannot be split");
                damage = Int32.Parse(UnsplitSource.Substring(0, spacePos));
                type = UnsplitSource.Substring(spacePos + 1);
            }
            public int Damage
            {
                get { return damage; }
                set { damage = value; }
            }
            public string Type
            {
                get { return type; }
                set { type = value; }
            }
        }
        #endregion

        void LoadSettings()
        {
            // Add items to the xmlSettings object here...
            xmlSettings.AddControlSetting(cbMultiDamageIsOne.Name, cbMultiDamageIsOne);
            xmlSettings.AddControlSetting(cbRecalcWardedHits.Name, cbRecalcWardedHits);
            xmlSettings.AddControlSetting(cbKatakana.Name, cbKatakana);
            xmlSettings.AddControlSetting(cbSParseConsider.Name, cbSParseConsider);
            xmlSettings.AddControlSetting(cbIncludeInterceptFocus.Name, cbIncludeInterceptFocus);

            if (File.Exists(settingsFile))
            {
                FileStream fs = new FileStream(settingsFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                XmlTextReader xReader = new XmlTextReader(fs);

                try
                {
                    while (xReader.Read())
                    {
                        if (xReader.NodeType == XmlNodeType.Element)
                        {
                            if (xReader.LocalName == "SettingsSerializer")
                            {
                                xmlSettings.ImportFromXml(xReader);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    lblStatus.Text = "Error loading settings: " + ex.Message;
                }
                xReader.Close();
            }
        }
        void SaveSettings()
        {
            FileStream fs = new FileStream(settingsFile, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
            XmlTextWriter xWriter = new XmlTextWriter(fs, Encoding.UTF8);
            xWriter.Formatting = Formatting.Indented;
            xWriter.Indentation = 1;
            xWriter.IndentChar = '\t';
            xWriter.WriteStartDocument(true);
            xWriter.WriteStartElement("Config");    // <Config>
            xWriter.WriteStartElement("SettingsSerializer");    // <Config><SettingsSerializer>
            xmlSettings.ExportToXml(xWriter);    // Fill the SettingsSerializer XML
            xWriter.WriteEndElement();    // </SettingsSerializer>
            xWriter.WriteEndElement();    // </Config>
            xWriter.WriteEndDocument();    // Tie up loose ends (shouldn't be any)
            xWriter.Flush();    // Flush the file buffer to disk
            xWriter.Close();
        }

        private void cbRecalcWardedHits_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Ward で軽減したダメージ分を加算して、本来の値を出力します。ただし、Stoneskin によって防いだ値は再計算できません。");
        }
        private void cbMultiDamageIsOne_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("有効にすると、複数属性攻撃(300 斬撃ダメージ、5 毒ダメージ、5 病気ダメージ)を合計ダメージ(「300/5/5 斬撃/毒/病気」を 「310」）で出力します。無効にすると、複数属性攻撃を複数回の攻撃として出力します。これは命中率の計算に影響します。");
        }
        private void cbKatakana_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("日本語と英語が併記されているものについて、日本語を有効にします。(例：ゾーンやアビリティ名など)");
        }
        private void cbSParseConsider_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("選択した一覧に /con, /whogroup, /whoraid コマンドでマークしたキャラクタを追加します。");
        }
        private void cbDebugLog_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("日本語プラグインで解析できなかったログを、指定ファイルに出力します。(デバッグ用機能です。普段はoffにしてください。動作が重くなります)");
        }
        private void cbIncludeInterceptFocus_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("チャネラーペットのフォーカスダメージを解析します。(攻撃者のDPSや命中率が正しく表示されなくなります。既に読み込んだデータには反映しません)");
        }

        private string GetIntCommas()
        {
            return ActGlobals.mainTableShowCommas ? "#,0" : "0";
        }
        private string GetFloatCommas()
        {
            return ActGlobals.mainTableShowCommas ? "#,0.00" : "0.00";
        }
        private void SetupEQ2EnglishEnvironment()
        {
            CultureInfo usCulture = new CultureInfo("en-US");    // This is for SQL syntax; do not change


            if (!CombatantData.ColumnDefs.ContainsKey("CritTypes"))
                CombatantData.ColumnDefs.Add("CritTypes", new CombatantData.ColumnDef("CritTypes", true, "VARCHAR(32)", "CritTypes", CombatantDataGetCritTypes, CombatantDataGetCritTypes, (Left, Right) => { return CombatantDataGetCritTypes(Left).CompareTo(CombatantDataGetCritTypes(Right)); }));
            if (!CombatantData.ExportVariables.ContainsKey("crittypes"))
                CombatantData.ExportVariables.Add("crittypes", new CombatantData.TextExportFormatter("crittypes", "Critical Types", "Distribution of Critical Types  (Normal|Legendary|Fabled|Mythical)", (Data, Extra) => { return CombatantFormatSwitch(Data, "crittypes", Extra); }));

            if (!DamageTypeData.ColumnDefs.ContainsKey("CritTypes"))
                DamageTypeData.ColumnDefs.Add("CritTypes", new DamageTypeData.ColumnDef("CritTypes", true, "VARCHAR(32)", "CritTypes", DamageTypeDataGetCritTypes, DamageTypeDataGetCritTypes));

            if (!AttackType.ColumnDefs.ContainsKey("CritTypes"))
                AttackType.ColumnDefs.Add("CritTypes", new AttackType.ColumnDef("CritTypes", true, "VARCHAR(32)", "CritTypes", AttackTypeGetCritTypes, AttackTypeGetCritTypes, (Left, Right) => { return AttackTypeGetCritTypes(Left).CompareTo(AttackTypeGetCritTypes(Right)); }));

            if (!MasterSwing.ColumnDefs.ContainsKey("CriticalStr"))
                MasterSwing.ColumnDefs.Add("CriticalStr", new MasterSwing.ColumnDef("CriticalStr", true, "VARCHAR(32)", "CriticalStr", (Data) =>
                {
                    if (Data.Tags.ContainsKey("CriticalStr"))
                        return (string)Data.Tags["CriticalStr"];
                    else
                        return "None";
                }, (Data) =>
                {
                    if (Data.Tags.ContainsKey("CriticalStr"))
                        return (string)Data.Tags["CriticalStr"];
                    else
                        return "None";
                }, (Left, Right) =>
                {
                    string left = Left.Tags.ContainsKey("CriticalStr") ? (string)Left.Tags["CriticalStr"] : "None";
                    string right = Right.Tags.ContainsKey("CriticalStr") ? (string)Right.Tags["CriticalStr"] : "None";
                    return left.CompareTo(right);
                }));

            foreach (KeyValuePair<string, MasterSwing.ColumnDef> pair in MasterSwing.ColumnDefs)
                pair.Value.GetCellForeColor = (Data) => { return GetSwingTypeColor(Data.SwingType); };

            ActGlobals.oFormActMain.ValidateLists();
            ActGlobals.oFormActMain.ValidateTableSetup();
        }
        private string CombatantDataGetCritTypes(CombatantData Data)
        {
            AttackType at;
            if (Data.AllOut.TryGetValue(ActGlobals.ActLocalization.LocalizationStrings["attackTypeTerm-all"].DisplayedText, out at))
            {
                return AttackTypeGetCritTypes(at);
            }
            else
                return "-";
        }
        private string DamageTypeDataGetCritTypes(DamageTypeData Data)
        {
            AttackType at;
            if (Data.Items.TryGetValue(ActGlobals.ActLocalization.LocalizationStrings["attackTypeTerm-all"].DisplayedText, out at))
            {
                return AttackTypeGetCritTypes(at);
            }
            else
                return "-";
        }
        private string AttackTypeGetCritTypes(AttackType Data)
        {
            int crit = 0;
            int lCrit = 0;
            int fCrit = 0;
            int mCrit = 0;
            for (int i = 0; i < Data.Items.Count; i++)
            {
                MasterSwing ms = Data.Items[i];
                if (ms.Critical)
                {
                    crit++;
                    if (!ms.Tags.ContainsKey("CriticalStr"))
                        continue;
                    if (((string)ms.Tags["CriticalStr"]).Contains("レジェンダリ"))
                    {
                        lCrit++;
                        continue;
                    }
                    if (((string)ms.Tags["CriticalStr"]).Contains("フェイブルド"))
                    {
                        fCrit++;
                        continue;
                    }
                    if (((string)ms.Tags["CriticalStr"]).Contains("ミシカル"))
                    {
                        mCrit++;
                        continue;
                    }
                }
            }
            float lCritPerc = ((float)lCrit / (float)crit) * 100f;
            float fCritPerc = ((float)fCrit / (float)crit) * 100f;
            float mCritPerc = ((float)mCrit / (float)crit) * 100f;
            if (crit == 0)
                return "-";
            return String.Format("{0:0.0}%L - {1:0.0}%F - {2:0.0}%M", lCritPerc, fCritPerc, mCritPerc);
        }
        private Color GetSwingTypeColor(int SwingType)
        {
            switch (SwingType)
            {
                case 1:
                case 2:
                    return Color.Crimson;
                case 3:
                    return Color.Blue;
                case 4:
                    return Color.DarkRed;
                case 5:
                    return Color.DarkOrange;
                case 8:
                    return Color.DarkOrchid;
                case 9:
                    return Color.DodgerBlue;
                default:
                    return Color.Black;
            }
        }
        private string EncounterFormatSwitch(EncounterData Data, List<CombatantData> SelectiveAllies, string VarName, string Extra)
        {
            long damage = 0;
            long healed = 0;
            int swings = 0;
            int hits = 0;
            int crits = 0;
            int heals = 0;
            int critheals = 0;
            int cures = 0;
            int misses = 0;
            int hitfail = 0;
            float tohit = 0;
            double dps = 0;
            double hps = 0;
            long healstaken = 0;
            long damagetaken = 0;
            long powerdrain = 0;
            long powerheal = 0;
            int kills = 0;
            int deaths = 0;

            switch (VarName)
            {
                case "maxheal":
                    return Data.GetMaxHeal(true, false);
                case "MAXHEAL":
                    return Data.GetMaxHeal(false, false);
                case "maxhealward":
                    return Data.GetMaxHeal(true, true);
                case "MAXHEALWARD":
                    return Data.GetMaxHeal(false, true);
                case "maxhit":
                    return Data.GetMaxHit(true);
                case "MAXHIT":
                    return Data.GetMaxHit(false);
                case "duration":
                    return Data.DurationS;
                case "DURATION":
                    return Data.Duration.TotalSeconds.ToString("0");
                case "damage":
                    foreach (CombatantData cd in SelectiveAllies)
                        damage += cd.Damage;
                    return damage.ToString();
                case "damage-m":
                    foreach (CombatantData cd in SelectiveAllies)
                        damage += cd.Damage;
                    return (damage / 1000000.0).ToString("0.00");
                case "DAMAGE-k":
                    foreach (CombatantData cd in SelectiveAllies)
                        damage += cd.Damage;
                    return (damage / 1000.0).ToString("0");
                case "DAMAGE-m":
                    foreach (CombatantData cd in SelectiveAllies)
                        damage += cd.Damage;
                    return (damage / 1000000.0).ToString("0");
                case "healed":
                    foreach (CombatantData cd in SelectiveAllies)
                        healed += cd.Healed;
                    return healed.ToString();
                case "swings":
                    foreach (CombatantData cd in SelectiveAllies)
                        swings += cd.Swings;
                    return swings.ToString();
                case "hits":
                    foreach (CombatantData cd in SelectiveAllies)
                        hits += cd.Hits;
                    return hits.ToString();
                case "crithits":
                    foreach (CombatantData cd in SelectiveAllies)
                        crits += cd.CritHits;
                    return crits.ToString();
                case "crithit%":
                    foreach (CombatantData cd in SelectiveAllies)
                        crits += cd.CritHits;
                    foreach (CombatantData cd in SelectiveAllies)
                        hits += cd.Hits;
                    float crithitperc = (float)crits / (float)hits;
                    return crithitperc.ToString("0'%");
                case "heals":
                    foreach (CombatantData cd in SelectiveAllies)
                        heals += cd.Heals;
                    return heals.ToString();
                case "critheals":
                    foreach (CombatantData cd in SelectiveAllies)
                        critheals += cd.CritHits;
                    return critheals.ToString();
                case "critheal%":
                    foreach (CombatantData cd in SelectiveAllies)
                        critheals += cd.CritHeals;
                    foreach (CombatantData cd in SelectiveAllies)
                        heals += cd.Heals;
                    float crithealperc = (float)critheals / (float)heals;
                    return crithealperc.ToString("0'%");
                case "cures":
                    foreach (CombatantData cd in SelectiveAllies)
                        cures += cd.CureDispels;
                    return cures.ToString();
                case "misses":
                    foreach (CombatantData cd in SelectiveAllies)
                        misses += cd.Misses;
                    return misses.ToString();
                case "hitfailed":
                    foreach (CombatantData cd in SelectiveAllies)
                        hitfail += cd.Blocked;
                    return hitfail.ToString();
                case "TOHIT":
                    foreach (CombatantData cd in SelectiveAllies)
                        tohit += cd.ToHit;
                    tohit /= SelectiveAllies.Count;
                    return tohit.ToString("0");
                case "DPS":
                case "ENCDPS":
                    foreach (CombatantData cd in SelectiveAllies)
                        damage += cd.Damage;
                    dps = damage / Data.Duration.TotalSeconds;
                    return dps.ToString("0");
                case "DPS-k":
                case "ENCDPS-k":
                    foreach (CombatantData cd in SelectiveAllies)
                        damage += cd.Damage;
                    dps = damage / Data.Duration.TotalSeconds;
                    return (dps / 1000.0).ToString("0");
                case "ENCHPS":
                    foreach (CombatantData cd in SelectiveAllies)
                        healed += cd.Healed;
                    hps = healed / Data.Duration.TotalSeconds;
                    return hps.ToString("0");
                case "ENCHPS-k":
                    foreach (CombatantData cd in SelectiveAllies)
                        healed += cd.Healed;
                    hps = healed / Data.Duration.TotalSeconds;
                    return (hps / 1000.0).ToString("0");
                case "tohit":
                    foreach (CombatantData cd in SelectiveAllies)
                        tohit += cd.ToHit;
                    tohit /= SelectiveAllies.Count;
                    return tohit.ToString("F");
                case "dps":
                case "encdps":
                    foreach (CombatantData cd in SelectiveAllies)
                        damage += cd.Damage;
                    dps = damage / Data.Duration.TotalSeconds;
                    return dps.ToString("F");
                case "dps-k":
                case "encdps-k":
                    foreach (CombatantData cd in SelectiveAllies)
                        damage += cd.Damage;
                    dps = damage / Data.Duration.TotalSeconds;
                    return (dps / 1000.0).ToString("F");
                case "enchps":
                    foreach (CombatantData cd in SelectiveAllies)
                        healed += cd.Healed;
                    hps = healed / Data.Duration.TotalSeconds;
                    return hps.ToString("F");
                case "enchps-k":
                    foreach (CombatantData cd in SelectiveAllies)
                        healed += cd.Healed;
                    hps = healed / Data.Duration.TotalSeconds;
                    return (hps / 1000.0).ToString("F");
                case "healstaken":
                    foreach (CombatantData cd in SelectiveAllies)
                        healstaken += cd.HealsTaken;
                    return healstaken.ToString();
                case "damagetaken":
                    foreach (CombatantData cd in SelectiveAllies)
                        damagetaken += cd.DamageTaken;
                    return damagetaken.ToString();
                case "powerdrain":
                    foreach (CombatantData cd in SelectiveAllies)
                        powerdrain += cd.PowerDamage;
                    return powerdrain.ToString();
                case "powerheal":
                    foreach (CombatantData cd in SelectiveAllies)
                        powerheal += cd.PowerReplenish;
                    return powerheal.ToString();
                case "kills":
                    foreach (CombatantData cd in SelectiveAllies)
                        kills += cd.Kills;
                    return kills.ToString();
                case "deaths":
                    foreach (CombatantData cd in SelectiveAllies)
                        deaths += cd.Deaths;
                    return deaths.ToString();
                case "title":
                    return Data.Title;

                default:
                    return VarName;
            }
        }
        private string CombatantFormatSwitch(CombatantData Data, string VarName, string Extra)
        {
            int len = 0;
            switch (VarName)
            {
                case "name":
                    return Data.Name;
                case "NAME":
                    len = Int32.Parse(Extra);
                    return Data.Name.Length - len > 0 ? Data.Name.Remove(len, Data.Name.Length - len).Trim() : Data.Name;
                case "NAME3":
                    len = 3;
                    return Data.Name.Length - len > 0 ? Data.Name.Remove(len, Data.Name.Length - len).Trim() : Data.Name;
                case "NAME4":
                    len = 4;
                    return Data.Name.Length - len > 0 ? Data.Name.Remove(len, Data.Name.Length - len).Trim() : Data.Name;
                case "NAME5":
                    len = 5;
                    return Data.Name.Length - len > 0 ? Data.Name.Remove(len, Data.Name.Length - len).Trim() : Data.Name;
                case "NAME6":
                    len = 6;
                    return Data.Name.Length - len > 0 ? Data.Name.Remove(len, Data.Name.Length - len).Trim() : Data.Name;
                case "NAME7":
                    len = 7;
                    return Data.Name.Length - len > 0 ? Data.Name.Remove(len, Data.Name.Length - len).Trim() : Data.Name;
                case "NAME8":
                    len = 8;
                    return Data.Name.Length - len > 0 ? Data.Name.Remove(len, Data.Name.Length - len).Trim() : Data.Name;
                case "NAME9":
                    len = 9;
                    return Data.Name.Length - len > 0 ? Data.Name.Remove(len, Data.Name.Length - len).Trim() : Data.Name;
                case "NAME10":
                    len = 10;
                    return Data.Name.Length - len > 0 ? Data.Name.Remove(len, Data.Name.Length - len).Trim() : Data.Name;
                case "NAME11":
                    len = 11;
                    return Data.Name.Length - len > 0 ? Data.Name.Remove(len, Data.Name.Length - len).Trim() : Data.Name;
                case "NAME12":
                    len = 12;
                    return Data.Name.Length - len > 0 ? Data.Name.Remove(len, Data.Name.Length - len).Trim() : Data.Name;
                case "NAME13":
                    len = 13;
                    return Data.Name.Length - len > 0 ? Data.Name.Remove(len, Data.Name.Length - len).Trim() : Data.Name;
                case "NAME14":
                    len = 14;
                    return Data.Name.Length - len > 0 ? Data.Name.Remove(len, Data.Name.Length - len).Trim() : Data.Name;
                case "NAME15":
                    len = 15;
                    return Data.Name.Length - len > 0 ? Data.Name.Remove(len, Data.Name.Length - len).Trim() : Data.Name;
                case "DURATION":
                    return Data.Duration.TotalSeconds.ToString("0");
                case "duration":
                    return Data.DurationS;
                case "maxhit":
                    return Data.GetMaxHit(true);
                case "MAXHIT":
                    return Data.GetMaxHit(false);
                case "maxheal":
                    return Data.GetMaxHeal(true, false);
                case "MAXHEAL":
                    return Data.GetMaxHeal(false, false);
                case "maxhealward":
                    return Data.GetMaxHeal(true, true);
                case "MAXHEALWARD":
                    return Data.GetMaxHeal(false, true);
                case "damage":
                    return Data.Damage.ToString();
                case "damage-k":
                    return (Data.Damage / 1000.0).ToString("0.00");
                case "damage-m":
                    return (Data.Damage / 1000000.0).ToString("0.00");
                case "DAMAGE-k":
                    return (Data.Damage / 1000.0).ToString("0");
                case "DAMAGE-m":
                    return (Data.Damage / 1000000.0).ToString("0");
                case "healed":
                    return Data.Healed.ToString();
                case "swings":
                    return Data.Swings.ToString();
                case "hits":
                    return Data.Hits.ToString();
                case "crithits":
                    return Data.CritHits.ToString();
                case "critheals":
                    return Data.CritHeals.ToString();
                case "crittypes":
                    return CombatantDataGetCritTypes(Data);
                case "crithit%":
                    return Data.CritDamPerc.ToString("0'%");
                case "critheal%":
                    return Data.CritHealPerc.ToString("0'%");
                case "heals":
                    return Data.Heals.ToString();
                case "cures":
                    return Data.CureDispels.ToString();
                case "misses":
                    return Data.Misses.ToString();
                case "hitfailed":
                    return Data.Blocked.ToString();
                case "TOHIT":
                    return Data.ToHit.ToString("0");
                case "DPS":
                    return Data.DPS.ToString("0");
                case "DPS-k":
                    return (Data.DPS / 1000.0).ToString("0");
                case "ENCDPS":
                    return Data.EncDPS.ToString("0");
                case "ENCDPS-k":
                    return (Data.EncDPS / 1000.0).ToString("0");
                case "ENCHPS":
                    return Data.EncHPS.ToString("0");
                case "ENCHPS-k":
                    return (Data.EncHPS / 1000.0).ToString("0");
                case "tohit":
                    return Data.ToHit.ToString("F");
                case "dps":
                    return Data.DPS.ToString("F");
                case "dps-k":
                    return (Data.DPS / 1000.0).ToString("F");
                case "encdps":
                    return Data.EncDPS.ToString("F");
                case "encdps-k":
                    return (Data.EncDPS / 1000.0).ToString("F");
                case "enchps":
                    return Data.EncHPS.ToString("F");
                case "enchps-k":
                    return (Data.EncHPS / 1000.0).ToString("F");
                case "healstaken":
                    return Data.HealsTaken.ToString();
                case "damagetaken":
                    return Data.DamageTaken.ToString();
                case "powerdrain":
                    return Data.PowerDamage.ToString();
                case "powerheal":
                    return Data.PowerReplenish.ToString();
                case "kills":
                    return Data.Kills.ToString();
                case "deaths":
                    return Data.Deaths.ToString();
                case "damage%":
                    return Data.DamagePercent;
                case "healed%":
                    return Data.HealedPercent;
                case "threatstr":
                    return Data.GetThreatStr("Threat (Out)");
                case "threatdelta":
                    return Data.GetThreatDelta("Threat (Out)").ToString();
                case "n":
                    return "\n";
                case "t":
                    return "\t";

                default:
                    return VarName;
            }
        }
        private string GetAttackTypeSwingType(AttackType Data)
        {
            int swingType = 100;
            List<int> swingTypes = new List<int>();
            List<MasterSwing> cachedItems = new List<MasterSwing>(Data.Items);
            for (int i = 0; i < cachedItems.Count; i++)
            {
                MasterSwing s = cachedItems[i];
                if (swingTypes.Contains(s.SwingType) == false)
                    swingTypes.Add(s.SwingType);
            }
            if (swingTypes.Count == 1)
                swingType = swingTypes[0];

            return swingType.ToString();
        }
        private string GetDamageTypeGrouping(DamageTypeData Data)
        {
            string grouping = string.Empty;

            int swingTypeIndex = 0;
            if (Data.Outgoing)
            {
                grouping += "attacker=" + Data.Parent.Name;
                foreach (KeyValuePair<int, List<string>> links in CombatantData.SwingTypeToDamageTypeDataLinksOutgoing)
                {
                    foreach (string damageTypeLabel in links.Value)
                    {
                        if (Data.Type == damageTypeLabel)
                        {
                            grouping += String.Format("&swingtype{0}={1}", swingTypeIndex++ == 0 ? string.Empty : swingTypeIndex.ToString(), links.Key);
                        }
                    }
                }
            }
            else
            {
                grouping += "victim=" + Data.Parent.Name;
                foreach (KeyValuePair<int, List<string>> links in CombatantData.SwingTypeToDamageTypeDataLinksIncoming)
                {
                    foreach (string damageTypeLabel in links.Value)
                    {
                        if (Data.Type == damageTypeLabel)
                        {
                            grouping += String.Format("&swingtype{0}={1}", swingTypeIndex++ == 0 ? string.Empty : swingTypeIndex.ToString(), links.Key);
                        }
                    }
                }
            }

            return grouping;
        }
    }
}
