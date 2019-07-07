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

[assembly: AssemblyTitle("Japanese Parsing Engine")]
[assembly: AssemblyDescription("Plugin based parsing Japanese EQ2 Sebilis server running the Japanese client")]
[assembly: AssemblyCompany("Gardin of Sebillis and Tzalik of Sebillis and Mayia of Sebilis")]
[assembly: AssemblyVersion("1.0.2.9")]

// NOTE: このpluginは、Tzalik様が公開していたpluginを元に改造したものです。（元バージョン配布サイト様：https://sites.google.com/site/eq2actjpn/home）
// NOTE: 解析者向け（＝自分用）に「pluginで解析できなかったログをファイルに出力する」隠し機能を搭載しております。ファイルの１行目の // を外すと利用可能。
// NOTE: レジェンダリ以上のクリティカルは、表示のみ対応しています。（"special"欄に表示されます）
// NOTE: チャネラーのログがうまく取り込めていなかった問題に対応いたしました。（まだあるかも・・・）
// NOTE: Obanburumai様のご協力により、EQ2側のバグによりアーツ名の直後で改行されていて取り込めなかったアーツ（ラッキー・ギャンビット、ワイルド・アクリーション、他にもあるかも？）を取得できるようになりました。
////////////////////////////////////////////////////////////////////////////////
// $Date: 2015-01-25 18:20:48 +0900 (2015/01/25 (日)) $
// $Rev: 21 $
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
            this.cbMultiDamageIsOne.Location = new System.Drawing.Point(13, 13);
            this.cbMultiDamageIsOne.Margin = new System.Windows.Forms.Padding(3, 1, 3, 1);
            this.cbMultiDamageIsOne.Name = "cbMultiDamageIsOne";
            this.cbMultiDamageIsOne.Size = new System.Drawing.Size(509, 16);
            this.cbMultiDamageIsOne.TabIndex = 5;
            this.cbMultiDamageIsOne.Text = "（※未対応です）複数属性ダメージを1回攻撃として記録する。(既に読み込んだデータには反映しません)";
            this.cbMultiDamageIsOne.MouseHover += new System.EventHandler(this.cbMultiDamageIsOne_MouseHover);
            // 
            // cbRecalcWardedHits
            // 
            this.cbRecalcWardedHits.AutoSize = true;
            this.cbRecalcWardedHits.Checked = true;
            this.cbRecalcWardedHits.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbRecalcWardedHits.Location = new System.Drawing.Point(13, 48);
            this.cbRecalcWardedHits.Margin = new System.Windows.Forms.Padding(3, 1, 3, 1);
            this.cbRecalcWardedHits.Name = "cbRecalcWardedHits";
            this.cbRecalcWardedHits.Size = new System.Drawing.Size(426, 16);
            this.cbRecalcWardedHits.TabIndex = 7;
            this.cbRecalcWardedHits.Text = "Ward で受けた値を本来の値で再計算する。(既に読み込んだデータには反映しません)";
            this.cbRecalcWardedHits.MouseHover += new System.EventHandler(this.cbRecalcWardedHits_MouseHover);
            // 
            // cbKatakana
            // 
            this.cbKatakana.AutoSize = true;
            this.cbKatakana.Checked = true;
            this.cbKatakana.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbKatakana.Location = new System.Drawing.Point(13, 30);
            this.cbKatakana.Margin = new System.Windows.Forms.Padding(3, 1, 3, 1);
            this.cbKatakana.Name = "cbKatakana";
            this.cbKatakana.Size = new System.Drawing.Size(321, 16);
            this.cbKatakana.TabIndex = 6;
            this.cbKatakana.Text = "表記を日本語にする。(既に読み込んだデータには反映しません)";
            this.cbKatakana.MouseHover += new System.EventHandler(this.cbKatakana_MouseHover);
            // 
            // cbSParseConsider
            // 
            this.cbSParseConsider.AutoSize = true;
            this.cbSParseConsider.Checked = true;
            this.cbSParseConsider.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbSParseConsider.Location = new System.Drawing.Point(13, 66);
            this.cbSParseConsider.Margin = new System.Windows.Forms.Padding(3, 1, 3, 1);
            this.cbSParseConsider.Name = "cbSParseConsider";
            this.cbSParseConsider.Size = new System.Drawing.Size(514, 16);
            this.cbSParseConsider.TabIndex = 8;
            this.cbSParseConsider.Text = "（※未対応です）選択した一覧に /con, /whogroup, /whoraid コマンドでマークしたキャラクタを追加する。";
            this.cbSParseConsider.MouseHover += new System.EventHandler(this.cbSParseConsider_MouseHover);
            #if DEBUG_FILE_OUTPUT
            // 
            // cbDebugLog
            // 
            this.cbDebugLog.AutoSize = true;
            this.cbDebugLog.Checked = false;
            this.cbDebugLog.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbDebugLog.Location = new System.Drawing.Point(14, 84);
            this.cbDebugLog.Margin = new System.Windows.Forms.Padding(3, 1, 3, 1);
            this.cbDebugLog.Name = "cbDebugLog";
            this.cbDebugLog.Size = new System.Drawing.Size(677, 16);
            this.cbDebugLog.TabIndex = 10;
            this.cbDebugLog.Text = "（※デバッグ用機能なので、普段はoffにしてください。動作が重くなります）日本語プラグインで解析できなかったログを指定ファイルに出力する。";
            this.cbDebugLog.MouseHover += new System.EventHandler(this.cbDebugLog_MouseHover);
            // 
            // tbDebugFileName
            // 
            this.tbDebugFileName.Location = new System.Drawing.Point(33, 97);
            this.tbDebugFileName.Name = "tbDebugFileName";
            this.tbDebugFileName.Size = new System.Drawing.Size(594, 19);
            this.tbDebugFileName.Text = @"D:\ACTJpnParser_debug.txt";
            this.tbDebugFileName.TabIndex = 11;
            // 
            // btDebugFileName
            // 
            this.btDebugFileName.Location = new System.Drawing.Point(626, 97);
            this.btDebugFileName.Name = "btDebugFileName";
            this.btDebugFileName.Size = new System.Drawing.Size(59, 21);
            this.btDebugFileName.TabIndex = 12;
            this.btDebugFileName.Text = "...";
            this.btDebugFileName.UseVisualStyleBackColor = true;
            this.btDebugFileName.Click += new System.EventHandler(this.btDebugFileName_Click);
            #endif
            // 
            // ACT_Jpn_Parser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            #if DEBUG_FILE_OUTPUT
            this.Controls.Add(this.btDebugFileName);
            this.Controls.Add(this.tbDebugFileName);
            this.Controls.Add(this.cbDebugLog);
            #endif
            this.Controls.Add(this.cbKatakana);
            this.Controls.Add(this.cbMultiDamageIsOne);
            this.Controls.Add(this.cbSParseConsider);
            this.Controls.Add(this.cbRecalcWardedHits);
            this.Name = "ACT_Jpn_Parser";
            this.Size = new System.Drawing.Size(694, 121);
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

            int dcIndex = -1;    // Find the Data Correction node in the Options tab
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
            ActGlobals.oFormActMain.BeforeLogLineRead += new LogLineEventDelegate(oFormActMain_BeforeLogLineRead);
            ActGlobals.oFormActMain.BeforeCombatAction += new CombatActionDelegate(oFormActMain_BeforeCombatAction);
            ActGlobals.oFormActMain.AfterCombatAction += new CombatActionDelegate(oFormActMain_AfterCombatAction);
            ActGlobals.oFormActMain.OnLogLineRead += new LogLineEventDelegate(oFormActMain_OnLogLineRead);
            lblStatus.Text = "Plugin は有効です。";
        }

        public void DeInitPlugin()
        {
            ActGlobals.oFormActMain.BeforeLogLineRead -= oFormActMain_BeforeLogLineRead;
            ActGlobals.oFormActMain.BeforeCombatAction -= oFormActMain_BeforeCombatAction;
            ActGlobals.oFormActMain.AfterCombatAction -= oFormActMain_AfterCombatAction;
            ActGlobals.oFormActMain.OnLogLineRead -= oFormActMain_OnLogLineRead;

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
        //Regex petSplit = new Regex(@"(?<petName>\w* ?)<(?<attacker>\w+)['?の](?<s>s?) (?<petClass>.+)>", RegexOptions.Compiled);
        Regex petSplit = new Regex(@"(?<attacker>\w+)(?:'s|の) ?(?<petName>.+)", RegexOptions.Compiled);
        //Regex engKillSplit = new Regex("(?<mob>.+?) in .+", RegexOptions.Compiled);
        Regex romanjiSplit = new Regex(@"\\r[iapsmq]:(?<katakana>.+?)\\::?(?<romanji>.+)\\/r", RegexOptions.Compiled);
        Regex regexConsider = new Regex(logTimeStampRegexStr + @".+?You consider (?<player>.+?)\.\.\. .+", RegexOptions.Compiled);
        Regex regexWhogroup = new Regex(logTimeStampRegexStr + @"(?<name>[^ ]+) Lvl \d+ .+", RegexOptions.Compiled);
        Regex regexWhoraid = new Regex(logTimeStampRegexStr + @"\[\d+ [^\]]+\] (?<name>[^ ]+) \([^\)]+\)", RegexOptions.Compiled);
        Regex regexLogTimeStamp = new Regex(logTimeStampRegexStr, RegexOptions.Compiled);
        CombatActionEventArgs lastDamage = null;

        // -- Special Thanks Obanburumai  Start --
        Regex regexSkillEol = new Regex(logTimeStampRegexStr + @"..*\\/r ?(?=\r?\n|$)", RegexOptions.Compiled);
        string skillEolLogline = string.Empty;
        bool isSkillEol = false;
        // -- Special Thanks Obanburumai  End --

        private void PopulateRegexArray()
        {
            regexArray = new Regex[22];
            regexArray[0]  = new Regex(logTimeStampRegexStr + @"(?<victim>.+?)(?:は|が) ?(?:(?<skillType>.+?)による攻撃を受け、)?(?<damageAndType>\d+(?:ポイント)?の.+?)ダメージを負(?:いまし|っ)た。(?:[（(](?<special>.+?)[（)])?", RegexOptions.Compiled);
            regexArray[1]  = new Regex(logTimeStampRegexStr + @"(?<attacker>あなた|.+?)(?:は|が)(?<victim>.+?)(?:に ?|をヒット。)(?<damageAndType>.+?)の?ダメージ(?:を負わせた|を負わせました)?。(?:[（(](?<special>.+?)[）)])?", RegexOptions.Compiled);
            regexArray[2]  = new Regex(logTimeStampRegexStr + @"(?<victim>.+?)(?:が|は) ?(?<attacker>.+?)(?:の放った|の|'s ?)(?<skill>.+?)により(?<damageAndType>\d+(?:ポイントの| +).+?)ダメージを受け(?:た|ました)。(?:[（(](?<special>.+)[）)])?", RegexOptions.Compiled);
            regexArray[3]  = new Regex(logTimeStampRegexStr + @"(?<attacker>.+?)(?:の放った|の|'s ?)(?<skill>.+?)により、(?<victim>.+?)(?:が|は)(?<damageAndType>\d+(?:ポイントの| +).+?)ダメージを受け(?:た|ました)。(?:[（(](?<special>.+)[）)])?", RegexOptions.Compiled);
            regexArray[4]  = new Regex(logTimeStampRegexStr + @"(?<attacker>.+?)(?:の|'s ?)(?<skill>.+?)で ?(?<damageAndType>\d+ポイントの.+?)ダメージを受けましｈ?た。(?:[（(](?<special>.+?)[）)])?", RegexOptions.Compiled);
            regexArray[5]  = new Regex(logTimeStampRegexStr + @"(?<healer>.+?) ?(?:の|は|'s ?)(?<skill>.+)(?:が|によって、)(?<victim>.+?)を(?:(?:ヒール|修復|リペア)してい(?:ます|る)|回復させました)(?<damage>\d+) ?ヒットポイントの?(?<crit>(?:フェイブルド|レジェンダリ|ミシカル)?クリティカル)?。", RegexOptions.Compiled);
            regexArray[6]  = new Regex(logTimeStampRegexStr + @"(?<attacker>.+?)(?:が|は ?)(?<victim>.+?)を(?<skill>.+?で攻撃|攻撃|ヒット)(?:。はずした|(?:しましたが|しようとしましたが)、失敗しました)。(?:[（(](?<special>.+?)[）)])?", RegexOptions.Compiled);
            regexArray[7]  = new Regex(logTimeStampRegexStr + @"(?<attacker>.+?)(?:が|は)(?<victim>.+?)を攻撃(?:。|しましたが、)(?<why>.+?)(?:がうまく妨害|によって妨げられま|はうまくかわしま)した。(?:[（(].*(?<special>ブロック|反撃|回避|受け流し|レジスト|反射|強打|カウンター).*[）)])?", RegexOptions.Compiled);
            regexArray[8]  = new Regex(logTimeStampRegexStr + @"(?<attacker>.+?)(?:が|は){1}(?<victim>.+?)を(?<skill>.+?)で攻撃(?:しましたが、|。)(?<why>.+?)(?:がうまく妨害|によって妨げられま|はうまくかわしま)した。(?:[（(].*(?<special>ブロック|反撃|回避|受け流し|レジスト|反射|強打|カウンター).*[）)])?", RegexOptions.Compiled);
            regexArray[9]  = new Regex(logTimeStampRegexStr + @"(?<victim>.+?)を倒した。", RegexOptions.Compiled);
            // NOTE: ペットの撃墜数を（極力）召喚主の手柄としてカウントしています。でもこれは正しい動きなのだろうか？
            //regexArray[10] = new Regex(logTimeStampRegexStr + @"(?<victim>.+?)が(?<attacker>.+?)に殺された……。", RegexOptions.Compiled);
            regexArray[10] = new Regex(logTimeStampRegexStr + @"(?<victim>.+?)が(?<attacker>.+?)(?:(?:の|'s ?).+?)?に殺された……。", RegexOptions.Compiled);
            regexArray[11] = new Regex(logTimeStampRegexStr + @"(?<attacker>.+?)に殺された……。", RegexOptions.Compiled);
            regexArray[12] = new Regex(logTimeStampRegexStr + @"Unknown command: 'act (.+)'", RegexOptions.Compiled);
            regexArray[13] = new Regex(logTimeStampRegexStr + @"(?:(?<attacker>[^\\].+?)(?:が|の))?(?:(?<skill>.+?) ?(?:で|が|により))?(?<victim>.+?)?(?:を攻撃し|に命中し|にヒットし|攻撃を受け)、(?:ポイントパワーを)?(?<damage>\d+)ポイント(?:パワーを)?消耗(?:させ|し)(?:た|ました)。?(?:[（(](?<special>.+?)[）)])?", RegexOptions.Compiled);
            regexArray[14] = new Regex(logTimeStampRegexStr + @"(?<victim>.+?)は(?<skill>.+?) ?によ(?:って|り)ポイントパワーを(?<damage>\d+)(?:ポイント)?消耗し(?:た|ました)。(?:[（(](?<special>.+?)[）)])?", RegexOptions.Compiled);
            regexArray[15] = new Regex(logTimeStampRegexStr + @"(?<victim>.+)に対する(?<damage>\d+) ?ポイントダメージを(?<attacker>あなた|.+?)(?:の|'s) ?(?<skillType>.+?)が吸収した。", RegexOptions.Compiled);
            regexArray[16] = new Regex(logTimeStampRegexStr + @"(?<skill>.+)は(?<damage>\d+) ?ポイントのダメージを吸収し、(?<victim>.+?)へのダメージを防いだ(?:。)?", RegexOptions.Compiled);
            regexArray[17] = new Regex(logTimeStampRegexStr + @"You have entered (?<zone>.+?)\.", RegexOptions.Compiled);
            regexArray[18] = new Regex(logTimeStampRegexStr + @"(?<healer>.+?)(?:が|は|の|'s ?)(?<skill>.+?)(?:が|で|によって) ?(?<victim>.+?)をリフレッシュしてい(?:る|ます)(?<damage>\d+) ?マナポイント(?:の)?(?<special>(?:フェイブルド|レジェンダリ|ミシカル)?クリティカル)?。", RegexOptions.Compiled);
            regexArray[19] = new Regex(logTimeStampRegexStr + @"(?:(?<owner>.+?|あなた)(?:の|'s))? ?(?<skillType>.+?)(?:が|で) ?(?<victim>.+?)に対する(?<target>.*)のヘイト(?:順位)?を ?(?<direction>増加|減少) ?(?<damage>\d+) ?(?<dirType>脅威レベル|position)の?(?<crit>(?:フェイブルド|レジェンダリ|ミシカル)?クリティカル)?。", RegexOptions.Compiled);
            regexArray[20] = new Regex(logTimeStampRegexStr + @"(?<attacker>.+?|あなた)(?:の|'s) ?(?<skillType>.+?)が ?(?:(?<victim>.+?|あなた))の(?<affliction>.+?)を(?<action>ディスペル|治療)しました。", RegexOptions.Compiled);
            regexArray[21] = new Regex(logTimeStampRegexStr + @"(?<healer>.+?)[はが] ?(?<attacker>.+?)から(?<victim>.+?)へのダメージを ?(?<damage>\d+) ?減らし(?:まし)?た。", RegexOptions.Compiled);
        }

        void oFormActMain_BeforeLogLineRead(bool isImport, LogLineEventArgs logInfo)
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

            if (regexSkillEol.IsMatch(logInfo.logLine))
            {
                this.isSkillEol = true;
                this.skillEolLogline = logInfo.logLine;
                this.skillEolLogline = this.skillEolLogline.Replace("\r", "").Replace("\n", "");
            }
            // -- Special Thanks Obanburumai  End --
            else
            {
                for (int i = 0; i < regexArray.Length; i++)
                {
                    if (regexArray[i].IsMatch(logInfo.logLine))
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
                                logInfo.detectedType = Color.Red.ToArgb();
                                break;
                            case 5:
                                logInfo.detectedType = Color.Blue.ToArgb();
                                break;
                            case 6:
                            case 7:
                            case 8:
                                logInfo.detectedType = Color.DarkRed.ToArgb();
                                break;
                            case 13:
                                logInfo.detectedType = Color.DarkOrchid.ToArgb();
                                break;
                            case 14:
                            case 15:
                            case 16:
                                logInfo.detectedType = Color.DodgerBlue.ToArgb();
                                break;
                            default:
                                logInfo.detectedType = Color.Black.ToArgb();
                                break;
                        }
                        LogExeJpn(i + 1, logInfo.logLine, isImport);
                        break;
                    }
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
        void oFormActMain_BeforeCombatAction(bool isImport, CombatActionEventArgs actionInfo)
        {
            // Riposte/kontert/
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
                if ((int)actionInfo.damage == (int)Dnum.Unknown && actionInfo.damage.DamageString.Contains("reflect"))
                {
                    if (actionInfo.theAttackType == lastDamage.theAttackType && actionInfo.victim == lastDamage.attacker)
                    {
                        //lastDamage.special = "reflect";    // Too late to take effect
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
                if (logInfo.logLine.Contains("/whoraid search results"))
                    captureWhoraid = true;
                if (logInfo.logLine.EndsWith("players found"))
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
                    if (outputName.StartsWith("{f}"))
                        outputName = outputName.Substring(3);
                    ActGlobals.oFormActMain.SelectiveListAdd(outputName);
                    if (!isImport)
                        System.Media.SystemSounds.Beep.Play();
                }
            }
        }
        
        private void LogExeJpn(int logMatched, string logLine, bool isImport) {
            // 追加する処理の分岐フラグ
            int NONE_DAMAGE = 0;
            int ADD_DAMAGE = 1;
            int SKIP_DAMAGE = 2;
            
            bool isSelfAttack = false;
            
            int branchFlag = NONE_DAMAGE;
            
            string attacker, victim, damage, skillType, why, special, damageType, crit;
            Regex rE = regexArray[logMatched - 1];
            int swingType = 0;
            bool critical = false;
            List<DamageAndType> damageAndTypeArr = new List<DamageAndType>();

            DateTime time = ActGlobals.oFormActMain.LastKnownTime;
            
            Dnum addCombatInDamage = null;
            
            int gts = ActGlobals.oFormActMain.GlobalTimeSorter;
            
            // 初期化
            attacker = string.Empty;
            victim = string.Empty;
            damage = string.Empty;
            skillType = string.Empty;
            why = string.Empty;
            special = string.Empty;
            damageType = string.Empty;
            crit = string.Empty;

            switch (logMatched) {
            #region Case 1 [unsourced skill attacks]
            case 1:
                branchFlag = ADD_DAMAGE;
                attacker = "不明";
                victim = rE.Replace(logLine, "$1");
                skillType = rE.Replace(logLine, "$2");
                damage = rE.Replace(logLine, "$3");
                special = rE.Replace(logLine, "$4");
                special = String.IsNullOrEmpty(special) ? "None" : special;
                crit = special;
                swingType = (int)SwingTypeEnum.NonMelee;
                if (skillType.Length == 0) {
                    skillType = "不明";
                }
                if (!ActGlobals.oFormActMain.InCombat && !isImport) {
                    ActGlobals.oFormSpellTimers.NotifySpell(attacker.ToLower(), skillType, victim.Contains("あなた"), victim.ToLower(), true);
                    branchFlag = NONE_DAMAGE;
                    break;
                }
                
                break;
            #endregion
            #region Case 2 [melee attacks by yourself]
            case 2:
                attacker = rE.Replace(logLine, "$1");
                victim = rE.Replace(logLine, "$2");
                damage = rE.Replace(logLine, "$3");
                special = rE.Replace(logLine, "$4");
                crit = special;
                special = special.Replace("クリティカルヒット・", string.Empty).Trim();
                special = special.Replace("クリティカル・", string.Empty).Trim();
                special = special.Replace("クリティカルヒット", string.Empty).Trim();
                special = special.Replace("クリティカル", string.Empty).Trim();
                if(special.Trim() == ""){
                    special = "None";
                }
                if (petSplit.IsMatch(attacker)) {
                    skillType = petSplit.Replace(attacker, "$2") + "の攻撃";
                    attacker = petSplit.Replace(attacker, "$1");
                    swingType = (int)SwingTypeEnum.NonMelee;
                } else {
                    swingType = (int)SwingTypeEnum.Melee;
                }
                isSelfAttack = true;
                if (ActGlobals.oFormActMain.SetEncounter(time, attacker, victim)) {
                    branchFlag = ADD_DAMAGE;
                }
                break;
            #endregion
            #region Case 3 [melee/non-melee attacks by expect yourself 3]
            case 3:
                victim = rE.Replace(logLine, "$1");
                attacker = rE.Replace(logLine, "$2");
                skillType = rE.Replace(logLine, "$3");
                damage = rE.Replace(logLine, "$4");
                special = rE.Replace(logLine, "$5");
                crit = special;
                special = special.Replace("クリティカルヒット・", string.Empty).Trim();
                special = special.Replace("クリティカル・", string.Empty).Trim();
                special = special.Replace("クリティカルヒット", string.Empty).Trim();
                special = special.Replace("クリティカル", string.Empty).Trim();
                if(special.Trim() == ""){
                    special = "None";
                }
                isSelfAttack = true;
                if (skillType == "攻撃") {
                    swingType = (int)SwingTypeEnum.Melee;
                    skillType = "";
                }else{
                    swingType = (int)SwingTypeEnum.NonMelee;
                }
                if (ActGlobals.oFormActMain.SetEncounter(time, attacker, victim)) {
                    branchFlag = ADD_DAMAGE;
                }
                break;
            #endregion
            #region Case 4 [melee/non-melee attacks by expect yourself 2]
            case 4:
                attacker = rE.Replace(logLine, "$1");
                skillType = rE.Replace(logLine, "$2");
                victim = rE.Replace(logLine, "$3");
                damage = rE.Replace(logLine, "$4");
                special = rE.Replace(logLine, "$5");
                crit = special;
                special = special.Replace("クリティカルヒット・", string.Empty).Trim();
                special = special.Replace("クリティカル・", string.Empty).Trim();
                special = special.Replace("クリティカルヒット", string.Empty).Trim();
                special = special.Replace("クリティカル", string.Empty).Trim();
                if(special.Trim() == ""){
                    special = "None";
                }
                isSelfAttack = true;
                if (skillType == "攻撃") {
                    swingType = (int)SwingTypeEnum.Melee;
                    skillType = "";
                }else{
                    swingType = (int)SwingTypeEnum.NonMelee;
                }
                if (ActGlobals.oFormActMain.SetEncounter(time, attacker, victim)) {
                    branchFlag = ADD_DAMAGE;
                }
                break;
            #endregion
            #region Case 5 [melee/non-melee attacks by expect yourself 3]
            case 5:
                attacker = rE.Replace(logLine, "$1");
                skillType = rE.Replace(logLine, "$2");
                damage = rE.Replace(logLine, "$3");
                special = rE.Replace(logLine, "$4");
                crit = special;
                special = special.Replace("クリティカルヒット・", string.Empty).Trim();
                special = special.Replace("クリティカル・", string.Empty).Trim();
                special = special.Replace("クリティカルヒット", string.Empty).Trim();
                special = special.Replace("クリティカル", string.Empty).Trim();
                if(special.Trim() == ""){
                    special = "None";
                }
                isSelfAttack = true;
                victim = "あなた";
                if (skillType == "攻撃") {
                    swingType = (int)SwingTypeEnum.Melee;
                    skillType = "";
                }else{
                    swingType = (int)SwingTypeEnum.NonMelee;
                }
                if (ActGlobals.oFormActMain.SetEncounter(time, attacker, victim)) {
                    branchFlag = ADD_DAMAGE;
                }
                break;
            #endregion
            #region Case 6 [healing]
            case 6:
                if (!ActGlobals.oFormActMain.InCombat) {
                    branchFlag = NONE_DAMAGE;
                    break;
                }
                branchFlag = SKIP_DAMAGE;
                attacker = rE.Replace(logLine, "$1");
                skillType = rE.Replace(logLine, "$2");
                victim = rE.Replace(logLine, "$3");
                damage = rE.Replace(logLine, "$4");
                special = rE.Replace(logLine, "$5");
                crit = special;
                special = special.Replace("クリティカルヒット・", string.Empty).Trim();
                special = special.Replace("クリティカル・", string.Empty).Trim();
                special = special.Replace("クリティカルヒット", string.Empty).Trim();
                special = special.Replace("クリティカル", string.Empty).Trim();
                if(special.Trim() == ""){
                    special = "None";
                }
                damageType = "Hitpoints";
                swingType = (int)SwingTypeEnum.Healing;
                if (attacker == "あなた" && logLine.Contains("自分を")) {
                    victim = attacker;
                }
                addCombatInDamage = Int32.Parse(damage);
                break;
            #endregion
            #region Case 7 [misses]
            case 7:
                attacker = rE.Replace(logLine, "$1");
                victim = rE.Replace(logLine, "$2");
                why = rE.Replace(logLine, "$3");
                special = rE.Replace(logLine, "$4");
                addCombatInDamage = Dnum.Miss;

                isSelfAttack = true;

                if ((why == "攻撃") && (!petSplit.IsMatch(attacker)))
                {
                    swingType = (int)SwingTypeEnum.Melee;
                    skillType = why.Trim();
                }
                else { // スキルmiss
                    swingType = (int)SwingTypeEnum.NonMelee;
                    skillType = (Regex.Split( why , "で攻撃" ))[0].Trim();
                    if (petSplit.IsMatch(attacker))
                    {
                        skillType = petSplit.Replace(attacker, "$2") + "の" + skillType;
                        attacker = petSplit.Replace(attacker, "$1");
                    }
                }

                if (ActGlobals.oFormActMain.SetEncounter(time, attacker, victim)) {
                    branchFlag = SKIP_DAMAGE;
                }
                break;
            #endregion
            #region Case 8 [melee misses by interfer]
            case 8:
                attacker = rE.Replace(logLine, "$1");
                victim = rE.Replace(logLine, "$2");
                why = rE.Replace(logLine, "$3");
                special = rE.Replace(logLine, "$4");
                crit = special;

                skillType = "攻撃";
                damageType = "melee";
                swingType = (int)SwingTypeEnum.Melee;
                
                if (petSplit.IsMatch(attacker))
                {
                    swingType = (int)SwingTypeEnum.NonMelee;
                    skillType = petSplit.Replace(attacker, "$2") + "の" + skillType;
                    attacker = petSplit.Replace(attacker, "$1");
                }

                why = why.Replace(victim, string.Empty);
                why = why.Trim() + " " + special;
                addCombatInDamage = new Dnum(Dnum.Unknown, why.Trim());
                isSelfAttack = true;
                if (ActGlobals.oFormActMain.SetEncounter(time, attacker, victim)) {
                    branchFlag = SKIP_DAMAGE;
                }
                break;
            #endregion
            #region Case 9 [non-melee misses by interfer]
            case 9:
                attacker = rE.Replace(logLine, "$1");
                victim = rE.Replace(logLine, "$2");
                skillType = rE.Replace(logLine, "$3");
                why = special = rE.Replace(logLine, "$4");
                special = rE.Replace(logLine, "$5");
                crit = special;
                damageType = "non-melee";
                swingType = (int)SwingTypeEnum.NonMelee;
                
                if (petSplit.IsMatch(attacker))
                {
                    skillType = petSplit.Replace(attacker, "$2") + "の" + skillType;
                    attacker = petSplit.Replace(attacker, "$1");
                }

                why = why.Replace(victim, string.Empty);
                why = why.Trim() + " " + special;
                addCombatInDamage = new Dnum(Dnum.Unknown, why.Trim());
                isSelfAttack = true;
                if (ActGlobals.oFormActMain.SetEncounter(time, attacker, victim)) {
                    branchFlag = SKIP_DAMAGE;
                }
                break;
            #endregion
            #region Case 10 [killing]
            case 10:
                if (!ActGlobals.oFormActMain.InCombat) {
                    branchFlag = NONE_DAMAGE;
                    break;
                }
                branchFlag = SKIP_DAMAGE;
                attacker = "あなた";
                victim = rE.Replace(logLine, "$1");
                swingType = (int)SwingTypeEnum.NonMelee;
                ActGlobals.oFormSpellTimers.RemoveTimerMods(victim);
                ActGlobals.oFormSpellTimers.DispellTimerMods(victim);
                special = "None";
                skillType = "Killing";
                addCombatInDamage = Dnum.Death;
                damageType = "Death";
                break;
            #endregion
            #region Case 11 [killed]
            case 11:
                if (!ActGlobals.oFormActMain.InCombat) {
                    branchFlag = NONE_DAMAGE;
                    break;
                }
                branchFlag = SKIP_DAMAGE;
                victim = rE.Replace(logLine, "$1");
                attacker = rE.Replace(logLine, "$2");
                swingType = (int)SwingTypeEnum.NonMelee;
                ActGlobals.oFormSpellTimers.RemoveTimerMods(victim);
                ActGlobals.oFormSpellTimers.DispellTimerMods(victim);
                special = "None";
                skillType = "Killing";
                addCombatInDamage = Dnum.Death;
                damageType = "Death";
                break;
            #endregion
            #region Case 12 [killing yourself]
            case 12:
                if (!ActGlobals.oFormActMain.InCombat) {
                    branchFlag = NONE_DAMAGE;
                    break;
                }
                branchFlag = SKIP_DAMAGE;
                victim = "あなた";
                attacker = rE.Replace(logLine, "$1");
                swingType = (int)SwingTypeEnum.NonMelee;
                ActGlobals.oFormSpellTimers.RemoveTimerMods(victim);
                ActGlobals.oFormSpellTimers.DispellTimerMods(victim);
                special = "None";
                skillType = "Killing";
                addCombatInDamage = Dnum.Death;
                damageType = "Death";
                break;
            #endregion
            #region Case 13 [act commands]
        case 13:
                branchFlag = NONE_DAMAGE;
                ActGlobals.oFormActMain.ActCommands(rE.Replace(logLine, "$1"));
                break;
            #endregion
            #region Case 14 [power drain 1]
            case 14:
                branchFlag = SKIP_DAMAGE;
                attacker = rE.Replace(logLine, "$1");
                skillType = rE.Replace(logLine, "$2");
                victim = rE.Replace(logLine, "$3");
                damage = rE.Replace(logLine, "$4");
                crit = rE.Replace(logLine, "$5");
                special = "None";
                swingType = (int)SwingTypeEnum.PowerDrain;
                isSelfAttack = true;

                if (attacker.Length == 0) {
                    attacker = "あなた";
                }
                if (victim.Length == 0) {
                    victim = "あなた";
                }
                if (skillType.Length == 0) {
                    skillType = "攻撃";
                }
                
                if (ActGlobals.oFormActMain.SetEncounter(time, attacker, victim)) {
                    if (CheckWardedHit(victim, time)) {
                        addCombatInDamage = new Dnum(Int32.Parse(damage) + lastWardAmount, String.Format("{0}/{1}", lastWardAmount, damage));
                        damageType = "warded/non-melee";
                        lastWardAmount = 0;
                    } else {
                        addCombatInDamage = Int32.Parse(damage);
                        damageType = "non-melee";
                    }
                }
                break;
            #endregion
            #region Case 15 [power drain 2]
            case 15:
                branchFlag = SKIP_DAMAGE;
                attacker = "不明";
                victim = rE.Replace(logLine, "$1");
                skillType = rE.Replace(logLine, "$2");
                damage = rE.Replace(logLine, "$3");
                crit = rE.Replace(logLine, "$4");
                special = "None";
                swingType = (int)SwingTypeEnum.PowerDrain;
                isSelfAttack = true;

                if (ActGlobals.oFormActMain.SetEncounter(time, attacker, victim)) {
                    if (CheckWardedHit(victim, time)) {
                        addCombatInDamage = new Dnum(Int32.Parse(damage) + lastWardAmount, String.Format("{0}/{1}", lastWardAmount, damage));
                        damageType = "warded/non-melee";
                        lastWardAmount = 0;
                    } else {
                        addCombatInDamage = Int32.Parse(damage);
                        damageType = "non-melee";
                    }
                }
                break;
            #endregion
            #region Case 16 [ward absorbtion]
            case 16:
                if (!ActGlobals.oFormActMain.InCombat) {
                    branchFlag = NONE_DAMAGE;
                    break;
                }
                branchFlag = SKIP_DAMAGE;
                victim = rE.Replace(logLine, "$1");
                damage = rE.Replace(logLine, "$2");
                attacker = rE.Replace(logLine, "$3");
                skillType = rE.Replace(logLine, "$4");
                swingType = (int)SwingTypeEnum.Healing;
                special = "None";
                damageType = "Absorption";
                addCombatInDamage = Int32.Parse(damage);
                
                if (CheckWardedHit(victim, time)) {
                    lastWardAmount += Int32.Parse(damage);
                } else {
                    lastWardAmount = Int32.Parse(damage);
                }
                lastWardedTarget = victim;
                lastWardTime = time;
                break;
            #endregion
            #region Case 17 [ward absorbtion your spell]
            case 17:
                if (!ActGlobals.oFormActMain.InCombat) {
                    branchFlag = NONE_DAMAGE;
                    break;
                }
                branchFlag = SKIP_DAMAGE;
                skillType = rE.Replace(logLine, "$1");
                damage = rE.Replace(logLine, "$2");
                victim = rE.Replace(logLine, "$3");
                
                attacker = "あなた";
                swingType = (int)SwingTypeEnum.Healing;
                special = "None";
                damageType = "Absorption";
                addCombatInDamage = Int32.Parse(damage);
                
                if (CheckWardedHit(victim, time)) {
                    lastWardAmount += Int32.Parse(damage);
                } else {
                    lastWardAmount = Int32.Parse(damage);
                }
                lastWardedTarget = victim;
                lastWardTime = time;
                break;
            #endregion
            #region Case 18 [zone change]
            case 18:
                branchFlag = NONE_DAMAGE;
                if (logLine.Contains(" combat by "))
                    break;
                string zoneName = rE.Replace(logLine, "$1").Trim();
                if(romanjiSplit.IsMatch(zoneName)){
                    zoneName = translateForMultiple(zoneName);
                }
                ActGlobals.oFormActMain.ChangeZone(zoneName);
                break;
            #endregion
            #region Case 19 [power healing]
            case 19:
                if (!ActGlobals.oFormActMain.InCombat) {
                    branchFlag = NONE_DAMAGE;
                    break;
                }
                branchFlag = SKIP_DAMAGE;
                attacker = rE.Replace(logLine, "$1");
                skillType = rE.Replace(logLine, "$2");
                victim = rE.Replace(logLine, "$3");
                damage = rE.Replace(logLine, "$4");
                special = rE.Replace(logLine, "$5");
                swingType = (int)SwingTypeEnum.PowerHealing;
                damageType = "Power";
                crit = special;
                special = special.Replace("クリティカルヒット・", string.Empty).Trim();
                special = special.Replace("クリティカル・", string.Empty).Trim();
                special = special.Replace("クリティカルヒット", string.Empty).Trim();
                special = special.Replace("クリティカル", string.Empty).Trim();
                if(special.Trim() == ""){
                    special = "None";
                }
                addCombatInDamage = Int32.Parse(damage);
                break;
            #endregion
            #region Case 20 [threat]
            case 20:
                branchFlag = NONE_DAMAGE;
                string owner = rE.Replace(logLine, "$1");
                skillType = rE.Replace(logLine, "$2");
                victim = rE.Replace(logLine, "$3");
                attacker = rE.Replace(logLine, "$4");
                string direction = rE.Replace(logLine, "$5");
                damage = rE.Replace(logLine, "$6");
                string dtype = rE.Replace(logLine, "$7");
                special = rE.Replace(logLine, "$8");
                swingType = (int)SwingTypeEnum.Threat;

                if (owner.Length == 0) {
                    owner = "あなた";
                }
                if (attacker.Contains("自分")||attacker.Contains("相手")) {
                    attacker = owner;
                }
                isSelfAttack = true;

                bool increase = (direction == "増加");
                crit = special;
                special = special.Replace("クリティカルヒット・", string.Empty).Trim();
                special = special.Replace("クリティカル・", string.Empty).Trim();
                special = special.Replace("クリティカルヒット", string.Empty).Trim();
                special = special.Replace("クリティカル", string.Empty).Trim();
                if(special.Trim() == ""){
                    special = "None";
                }

                Dnum dDamage;
                bool positionChange = (dtype == "position");
                if (positionChange) {
                    dDamage = new Dnum(Dnum.ThreatPosition, String.Format("{0} Positions", Int32.Parse(damage)));
                }
                else {
                    dDamage = new Dnum(Int32.Parse(damage));
                }

                direction = increase ? "Increase" : "Decrease";

                if (ActGlobals.oFormActMain.SetEncounter(time, attacker, victim) || ActGlobals.oFormActMain.SetEncounter(time, owner, victim)) {
                    branchFlag = SKIP_DAMAGE;
                    damageType = direction;
                    addCombatInDamage = dDamage;
                }
                break;
            #endregion
            #region Case 21 [dispell/cure]
            case 21:
                branchFlag = NONE_DAMAGE;
                attacker = rE.Replace(logLine, "$1");
                skillType = rE.Replace(logLine, "$2");
                victim = rE.Replace(logLine, "$3");
                string attackType = rE.Replace(logLine, "$4");
                direction = rE.Replace(logLine, "$5");
                swingType = (int)SwingTypeEnum.CureDispel;

                if (attackType.Contains("Traumatic Swipe") || attackType.Contains("トラウマティック・スワイプ")){
                    ActGlobals.oFormSpellTimers.DispellTimerMods(victim);
                }

                bool cont = false;
                if (direction.Contains("治療")) {
                    cont = ActGlobals.oFormActMain.InCombat;
                } else {
                    cont = ActGlobals.oFormActMain.SetEncounter(time, attacker, victim);
                }
                if (cont) {
                    branchFlag = SKIP_DAMAGE;
                    special = attackType;
                    addCombatInDamage = 1;
                    damageType = direction;
                }
                break;
            #endregion
            #region Case 22 [damage interception]
                case 22:
                branchFlag = NONE_DAMAGE;
                attacker = rE.Replace(logLine, "$1");   // Inteceptor
                special = rE.Replace(logLine, "$2");    // Attacker
                victim = rE.Replace(logLine, "$3");     // Target
                damage = rE.Replace(logLine, "$4");     // Amount

                swingType = (int)SwingTypeEnum.Healing;
                skillType = "Channeler Pet";

                if (ActGlobals.oFormActMain.SetEncounter(time, attacker, victim)) {
                    branchFlag = SKIP_DAMAGE;
                    damageType = "Interception";
                    addCombatInDamage = Int32.Parse(damage);
                }
                break;
            #endregion
            default:
                branchFlag = NONE_DAMAGE;
                break;
            }
            
            
            if (attacker.Contains("あなた")){
                attacker = ActGlobals.charName;
            }
            if (victim.Contains("あなた") || victim.Contains("自分")){
                victim = ActGlobals.charName;
            }
            if(!critical){
                critical = crit.Contains("クリティカル");
            }
            
            if (isSelfAttack && (attacker == victim || attacker == petSplit.Replace(victim, "$2"))) {
                branchFlag = NONE_DAMAGE;
            }
            
            if(branchFlag == ADD_DAMAGE){
                damageAndTypeArr = EngGetDamageAndTypeArr(damage);
                AddDamageAttack(swingType, critical, special, attacker, skillType, damageAndTypeArr, time, gts, victim);
            }else if(branchFlag == SKIP_DAMAGE){
                AddCombatActionTrans(swingType, critical, special, attacker, skillType, addCombatInDamage, time, gts, victim, damageType);
            }
        }
        private bool CheckWardedHit(string victim, DateTime time)
        {
            return cbRecalcWardedHits.Checked && lastWardTime == time && lastWardedTarget == victim && lastWardAmount > 0;
        }
        private void AddDamageAttack(int swingType, bool critical, string special, string attacker, string skillType, List<DamageAndType> damageAndTypeArr, DateTime time, int gts, string victim)
        {
            int damageTotal = 0;
            if (cbMultiDamageIsOne.Checked)
            {
                string damageStr = string.Empty;
                string typeStr = string.Empty;
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
                bool nullSkillType = false;
                if (String.IsNullOrEmpty(skillType))
                    nullSkillType = true;
                for (int i = 0; i < damageAndTypeArr.Count; i++)
                {
                    damageTotal += damageAndTypeArr[i].Damage;
                    if (nullSkillType)
                        skillType = damageAndTypeArr[i].Type;
                    if (i == damageAndTypeArr.Count - 1 && CheckWardedHit(victim, time))
                    {
                        damageTotal += lastWardAmount;
                        lastWardAmount = 0;
                    }
                    AddCombatActionTrans(swingType, critical, special, attacker, skillType, damageTotal, time, gts, victim, damageAndTypeArr[i].Type);
                }
            }
        }
        public void AddCombatActionTrans(int SwingType, bool Critical, string Special, string Attacker, string theAttackType, Dnum Damage, DateTime Time, int TimeSorter, string Victim, string theDamageType)
        {
            if(romanjiSplit.IsMatch(theDamageType)){
                theDamageType = translateForMultiple(theDamageType);
            }
            if(romanjiSplit.IsMatch(theAttackType)){
                theAttackType = translateForMultiple(theAttackType);
            }
            if(romanjiSplit.IsMatch(Attacker)){
                Attacker = translateForMultiple(Attacker);
            }
            if(romanjiSplit.IsMatch(Victim)){
                Victim = translateForMultiple(Victim);
            }
            if(romanjiSplit.IsMatch(Special)){
                Special = translateForMultiple(Special);
            }
            
            ActGlobals.oFormActMain.AddCombatAction(SwingType, Critical, Special, Attacker, theAttackType, Damage, Time, TimeSorter, Victim, theDamageType);
        }
        private string translateForMultiple(string transTarget) {
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
        private List<DamageAndType> EngGetDamageAndTypeArr(string damageAndType) {
            List<DamageAndType> outList = new List<DamageAndType>();
            damageAndType = damageAndType.Replace(" and ", ", ");
            damageAndType = damageAndType.Replace("ポイントの", " ");
            string[] entries = damageAndType.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);
            

            for (int i = 0; i < entries.Length; i++)
            {
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
            ActGlobals.oFormActMain.SetOptionsHelpText("未対応です：複数属性攻撃(300 斬撃ダメージ、5 毒ダメージ、5 病気ダメージ)を合計ダメージ(「300/5/5 斬撃/毒/病気」を 「310」）で出力します。If disabled, each damage type will show up as an individual swing, IE three attacks: 300 crushing; 5 poison; 5 disease.  Having a single attack show up as multiple will have consequences when calculating ToHit%.");
        }
        private void cbKatakana_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("日本語と英語の表記があるものについて、日本語を有効にします。(例：ゾーンやアビリティ名など)");
        }
        private void cbSParseConsider_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("未対応です：The /con command simply adds some text to the log about your target's con-level.  The /whogroup and /whoraid commands will list the members of your group/raid respectively.  Using this option will allow you to quickly add players to the Selective Parsing list.");
        }

        #if DEBUG_FILE_OUTPUT
        private void cbDebugLog_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("（※デバッグ用機能です）pluginで解析できなかったログを、指定ファイルに出力します。動作が重くなるので普段はoffにしてください。");
        }

        private void btDebugFileName_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (this.tbDebugFileName.Text.Length == 0)
                ofd.FileName = "ACTJpnParser_debug.txt";
            else ofd.FileName = System.IO.Path.GetFileName(this.tbDebugFileName.Text);
            if (this.tbDebugFileName.Text.Length == 0)
                ofd.InitialDirectory = @"D:\";
            else ofd.FileName = System.IO.Path.GetDirectoryName(this.tbDebugFileName.Text);
            ofd.Filter = "テキストファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*";
            ofd.FilterIndex = 2;
            ofd.Title = "出力ファイル名を選択してください";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                this.tbDebugFileName.Text = ofd.FileName;
            }
        }
        #endif
    }
}

