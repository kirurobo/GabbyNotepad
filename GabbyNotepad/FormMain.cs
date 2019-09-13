using System;
using System.IO;
using System.Windows.Forms;
using System.Speech.Synthesis;

namespace GabbyNotepad
{
    public partial class FormMain : Form
    {
        /// <summary>
        /// 設定ファイルのパス
        /// </summary>
        static string SettingsPath = System.IO.Directory.GetParent(System.Reflection.Assembly.GetExecutingAssembly().Location).FullName + @"\config.json";

        Settings MySettings = new Settings();

        SpeechSynthesizer Synth = new SpeechSynthesizer();

        bool IsTextChanged = false;
        string FilePath = "";


        void InitializeSynthesizer()
        {
            Synth.SetOutputToDefaultAudioDevice();

            // 音声一覧の選択肢を準備
            ToolStripItemCollection collection = voiceToolStripMenuItem.DropDownItems;
            collection.Clear();

            var voiceList = Synth.GetInstalledVoices();
            foreach (var voice in voiceList)
            {
                ToolStripMenuItem item = new ToolStripMenuItem(voice.VoiceInfo.Name);
                collection.Add(item);

                if (voice.VoiceInfo.Id == Synth.Voice.Id)
                {
                    item.Checked = true;
                }
                item.Click += new EventHandler(voiceToolStripMenuItem_Click);
            }

        }

        public FormMain()
        {
            InitializeComponent();

            // 設定ファイル読み込み
            MySettings.Load(SettingsPath);
            ApplySettings();

            // 音声の準備
            InitializeSynthesizer();
        }

        /// <summary>
        /// 設定を反映
        /// </summary>
        private void ApplySettings()
        {
            speakToolStripMenuItem.Checked = MySettings.EnableWordByWord;
            textBoxMain.WordWrap = wordWrapToolStripMenuItem.Checked = MySettings.EnableWordWrap;
            
            if (MySettings.Font != null)
            {
                textBoxMain.Font = MySettings.Font;
            }

            try
            {
                if (string.IsNullOrEmpty(MySettings.VoiceName)) {
                    // 指定がなければ英語を探す
                    Synth.SelectVoiceByHints(
                        VoiceGender.NotSet,
                        VoiceAge.NotSet,
                        0,
                        new System.Globalization.CultureInfo("en-US", false)
                        );
                }
                else
                {
                    Synth.SelectVoice(MySettings.VoiceName);
                }
            }
            catch
            {
            }
        }

        /// <summary>
        /// 設定を保存
        /// </summary>
        private void SaveSettings()
        {
            MySettings.EnableWordByWord = speakToolStripMenuItem.Checked;
            MySettings.EnableWordWrap = textBoxMain.WordWrap;
            MySettings.Font = textBoxMain.Font;
            MySettings.VoiceName = Synth.Voice.Name;

            MySettings.Save(SettingsPath);
        }

        /// <summary>
        /// 入力時の処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBoxMain_KeyDown(object sender, KeyEventArgs e)
        {
            // 読み上げオンなら単語読み上げ処理
            if (speakToolStripMenuItem.Checked)
            {
                if (e.KeyCode == Keys.Space || e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
                {
                    string word = "";

                    // カーソル位置取得
                    int curpos = textBoxMain.SelectionStart;
                    if (curpos >= textBoxMain.Text.Length) curpos = textBoxMain.Text.Length - 1;

                    char chr = textBoxMain.Text[curpos];
                    // もし直前が文末らしければ、文全体を読み上げ
                    if (chr == '.' || chr == '!' || chr == '?')
                    {
                        var pos = FindSentence();
                        word = textBoxMain.Text.Substring(pos.Item1, pos.Item2);
                    }
                    else
                    {
                        // カーソル位置の手前にある単語を取得
                        while (curpos >= 0)
                        {
                            chr = textBoxMain.Text[curpos];
                            if (char.IsWhiteSpace(chr)) break;
                            word = chr + word;
                            curpos--;
                        }
                    }

                    // カーソル前に単語があれば読み上げ
                    if (word.Length > 0)
                    {
                        Synth.SpeakAsync(word);
                    }
                }
            }
        }


        /// <summary>
        /// 折り返しの設定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void wordWrapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMain.WordWrap = wordWrapToolStripMenuItem.Checked = !wordWrapToolStripMenuItem.Checked;
        }

        /// <summary>
        /// すべて選択
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMain.SelectAll();
        }

        /// <summary>
        /// 終了
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        /// <summary>
        /// 新規文書作成
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 編集破棄確認
            if (!ConfirmDiscardChanges()) return;

            // 新規テキスト
            textBoxMain.Text = "";
            textBoxMain.Select(0, 0);
            IsTextChanged = false;
            FilePath = "";
        }

        /// <summary>
        /// 編集が行われたことのチェック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBoxMain_TextChanged(object sender, EventArgs e)
        {
            IsTextChanged = true;
        }

        /// <summary>
        /// 破棄の確認
        /// </summary>
        /// <returns></returns>
        private bool ConfirmDiscardChanges()
        {
            // 編集破棄の確認
            if (IsTextChanged)
            {
                DialogResult res = MessageBox.Show(this, "Do you want to discard unsaved text?", Application.ProductName, MessageBoxButtons.OKCancel);
                if (res != DialogResult.OK) return false;
            }
            return true;
        }

        /// <summary>
        /// フォント選択
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void fontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FontDialog fontDialog = new FontDialog();
            fontDialog.Font = textBoxMain.Font;

            if (fontDialog.ShowDialog(this) == DialogResult.OK)
            {
                textBoxMain.Font = fontDialog.Font;
            }
        }

        /// <summary>
        /// 切り取り
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMain.Cut();
        }

        /// <summary>
        /// コピー
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMain.Copy();
        }

        /// <summary>
        /// 貼付け
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMain.Paste();
        }

        /// <summary>
        /// 段落を発声
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void speakParagraphToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (textBoxMain.SelectionLength == 0)
            {
                SelectSentence();
            }
            
            // カーソル前に単語があれば読み上げ
            if (textBoxMain.SelectionLength > 0)
            {
                Synth.SpeakAsync(textBoxMain.SelectedText);
            }
        }

        /// <summary>
        /// 現在一直前の一文を探して選択
        /// </summary>
        private void SelectSentence()
        {
            var position = FindSentence();

            // 文字列を選択
            textBoxMain.Select(position.Item1, position.Item2);
        }

        /// <summary>
        /// 文の範囲を自動選択
        /// </summary>
        /// <returns>(始点, 文字数)</returns>
        private Tuple<int, int> FindSentence()
        {
            // 行頭を探す
            int start = textBoxMain.SelectionStart;
            if (start >= textBoxMain.Text.Length) start = textBoxMain.Text.Length - 1;
            bool wordFlag = false;
            for (int i = start; i >= 0; i--)
            {
                char chr = textBoxMain.Text[i];
                if (chr == '.' || chr == '!' || chr == '?' || chr == '\n')
                {
                    if (wordFlag) break;
                    else continue;
                }

                if (char.IsWhiteSpace(chr)) continue;

                start = i;
                wordFlag = true;
            }

            // 行末を探す
            int end = textBoxMain.SelectionStart + textBoxMain.SelectionLength;
            if (end >= textBoxMain.TextLength) end = textBoxMain.TextLength - 1;
            for (int i = end; i < textBoxMain.TextLength; i++)
            {
                end = i;

                char chr = textBoxMain.Text[i];
                if (chr == '.' || chr == '!' || chr == '?' || chr == '\n') break;
            }

            return Tuple.Create(start, end - start + 1);
        }

        /// <summary>
        /// 全て発声
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void speakAllTextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Synth.SpeakAsync(textBoxMain.Text);
        }

        /// <summary>
        /// ファイルを開く
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 編集破棄確認
            if (!ConfirmDiscardChanges()) return;

            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.DefaultExt = ".txt";
            fileDialog.Filter = "Text files(*.txt)|*.txt|All files(*.*)|*.*";
            fileDialog.CheckFileExists = true;

            if (fileDialog.ShowDialog(this) == DialogResult.OK)
            {
                using (StreamReader reader = new StreamReader(fileDialog.FileName))
                {
                    textBoxMain.Text = reader.ReadToEnd();
                    reader.Close();

                    textBoxMain.Select(0, 0);

                    FilePath = fileDialog.FileName;
                    IsTextChanged = false;
                }
            }
        }

        /// <summary>
        /// ファイルを保存
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (FilePath == "")
            {
                saveAsToolStripMenuItem_Click(sender, e);
                return;
            }

            using (StreamWriter writer = File.CreateText(FilePath))
            {
                writer.Write(textBoxMain.Text);
                writer.Close();

                IsTextChanged = false;
            }
        }

        /// <summary>
        /// 名前を付けてファイルを保存
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog fileDialog = new SaveFileDialog();
            fileDialog.DefaultExt = ".txt";
            fileDialog.Filter = "Text files(*.txt)|*.txt";
            fileDialog.OverwritePrompt = true;

            if (fileDialog.ShowDialog(this) == DialogResult.OK)
            {
                using (StreamWriter writer = File.CreateText(fileDialog.FileName))
                {
                    writer.Write(textBoxMain.Text);
                    writer.Close();

                    FilePath = fileDialog.FileName;
                    IsTextChanged = false;
                }
            }
        }

        /// <summary>
        /// 発声停止
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void stopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Synth.SpeakAsyncCancelAll();
        }

        /// <summary>
        /// 音声の選択
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void voiceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int index = 0;
            foreach (ToolStripMenuItem item in voiceToolStripMenuItem.DropDownItems)
            {
                if (sender.Equals(item))
                {
                    item.Checked = true;
                    Synth.SelectVoice(item.Text);
                }
                else
                {
                    item.Checked = false;
                }
                index++;
            }
        }

        /// <summary>
        /// 終了前の処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();

            // 編集破棄確認
            if (!ConfirmDiscardChanges())
            {
                e.Cancel = true;
                return;
            }
        }
    }
}
