using System;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Text;

namespace GabbyNotepad
{
    public partial class FormMain : Form
    {
        /// <summary>
        /// 設定ファイルのパス
        /// 実行ファイルと同じディレクトリの config.json とする
        /// </summary>
        static readonly string SettingsPath = Path.Combine(
            Directory.GetParent(System.Reflection.Assembly.GetExecutingAssembly().Location).FullName,
            "config.json"
            );

        /// <summary>
        /// 文末となる文字の一覧
        /// </summary>
        static readonly char[] SentenceDelimiters = {
            '.', '!', '?', ':', ';', '\n', '\r', '\u0085','\u2028', '\u2029',
            '。', '．', '：', '；'
        };

        /// <summary>
        /// 行区切りとなる文字の一覧
        /// </summary>
        static readonly char[] LineDelimiters =
        {
            '\n', '\r', '\u0085', '\u2028', '\u2029'
        };

        /// <summary>
        /// ホイールで変更できるフォントサイズ
        /// </summary>
        static readonly float[] FontSizeArray =
        {
            8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72, 100, 150, 200, 400
        };

        Settings MySettings = new Settings();
        SpeechSynthesizer MySynthesizer = new SpeechSynthesizer();
        AboutBoxVersion MyAboutBox = new AboutBoxVersion();

        bool IsTextChanged = false;
        string CurrentFilePath = "";

        DateTime lastFontResizedTime = DateTime.Now;


        public FormMain()
        {
            InitializeComponent();

            // [Ctrl]+ホイールでのフォントサイズ変更
            MouseWheel += FormMain_MouseWheel;

            // 設定ファイル読み込み
            MySettings.Load(SettingsPath);
            ApplySettings();

            // 音声の準備
            InitializeSynthesizer();
        }

        /// <summary>
        /// 音声合成エンジンの初期化
        /// </summary>
        void InitializeSynthesizer()
        {
            MySynthesizer.SetOutputToDefaultAudioDevice();

            // メニューでの音声一覧の選択肢を準備
            ToolStripItemCollection collection = voiceToolStripMenuItem.DropDownItems;
            collection.Clear();
            var voiceList = MySynthesizer.GetInstalledVoices();
            foreach (var voice in voiceList)
            {
                ToolStripMenuItem item = new ToolStripMenuItem(voice.VoiceInfo.Name);
                collection.Add(item);

                if (voice.VoiceInfo.Id == MySynthesizer.Voice.Id)
                {
                    item.Checked = true;
                }
                item.Click += new EventHandler(voiceToolStripMenuItem_Click);
            }
        }

        /// <summary>
        /// 設定を反映
        /// </summary>
        private void ApplySettings()
        {
            // 逐次発声有無
            speakToolStripMenuItem.Checked = MySettings.EnableWordByWord;

            // 行の折り返し有無
            textBoxMain.WordWrap = wordWrapToolStripMenuItem.Checked = MySettings.EnableWordWrap;

            // フォント指定
            if (MySettings.Font != null)
            {
                textBoxMain.Font = MySettings.Font;
            }

            // 音声の選択
            try
            {
                if (string.IsNullOrEmpty(MySettings.VoiceName)) {
                    // 音声の指定がなければ、英語のものを探して選択
                    MySynthesizer.SelectVoiceByHints(
                        VoiceGender.NotSet,
                        VoiceAge.NotSet,
                        0,
                        new System.Globalization.CultureInfo("en-US", false)
                        );
                }
                else
                {
                    // 保存されていた設定の音声を選択
                    MySynthesizer.SelectVoice(MySettings.VoiceName);
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
            MySettings.VoiceName = MySynthesizer.Voice.Name;

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
                    int curpos = textBoxMain.SelectionStart - 1;
                    if (curpos >= textBoxMain.Text.Length) curpos = textBoxMain.Text.Length - 1;

                    // もし直前が文末らしければ文全体を読み上げ
                    if ((curpos >= 0) && (SentenceDelimiters.Contains(textBoxMain.Text[curpos])))
                    {
                        // 直前が改行でなければ、文全体を読み上げ
                        char chr = textBoxMain.Text[curpos];
                        if (chr != '\r' && chr != '\n')
                        {
                            var pos = FindSentence();
                            word = textBoxMain.Text.Substring(pos.Item1, pos.Item2);
                        }
                    }
                    else
                    {
                        // カーソル位置の手前にある単語を取得
                        while (curpos >= 0)
                        {
                            char chr = textBoxMain.Text[curpos];
                            if (char.IsWhiteSpace(chr)) break;
                            word = chr + word;
                            curpos--;
                        }
                    }

                    // カーソル前に単語があれば読み上げ
                    if (word.Length > 0)
                    {
                        MySynthesizer.SpeakAsync(word);
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
            CurrentFilePath = "";
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
                MySynthesizer.SpeakAsync(textBoxMain.SelectedText);
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
            // 現在カーソル位置（または選択範囲）を取得
            int start = textBoxMain.SelectionStart - 1;
            if (start >= textBoxMain.TextLength) start = textBoxMain.TextLength - 1;

            bool isWordFound = false;  // 発声すべき文字が見つかればtrue

            // 一文字ずつ戻って行頭を探す
            for (int i = start; i >= 0; i--)
            {
                char chr = textBoxMain.Text[i];

                // 単語区切りにあたったかの検査
                if (SentenceDelimiters.Contains(chr))
                {
                    if (isWordFound) break; // 読み上げ対象の単語が見つかっていれば終了
                    else continue;
                }

                if (char.IsWhiteSpace(chr)) continue;

                start = i;
                isWordFound = true;
            }
            if (start < 0) start = 0;

            // 行末を探す
            int end = textBoxMain.SelectionStart - 1 + textBoxMain.SelectionLength;
            if (end < 0) end = 0;
            if (end >= textBoxMain.TextLength) end = textBoxMain.TextLength - 1;
            for (int i = end; i < textBoxMain.TextLength; i++)
            {
                end = i;

                char chr = textBoxMain.Text[i];
                if (SentenceDelimiters.Contains(chr)) break;
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
            MySynthesizer.SpeakAsync(textBoxMain.Text);
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

                    CurrentFilePath = fileDialog.FileName;
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
            if (CurrentFilePath == "")
            {
                saveAsToolStripMenuItem_Click(sender, e);
                return;
            }

            using (StreamWriter writer = File.CreateText(CurrentFilePath))
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

                    CurrentFilePath = fileDialog.FileName;
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
            MySynthesizer.SpeakAsyncCancelAll();
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
                    MySynthesizer.SelectVoice(item.Text);
                }
                else
                {
                    item.Checked = false;
                }
                index++;
            }
        }

        /// <summary>
        /// バージョン情報
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MyAboutBox.Synthesizer = MySynthesizer;
            MyAboutBox.ShowDialog(this);
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

        /// <summary>
        /// [Ctrl] + ホイールでフォントサイズ変更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormMain_MouseWheel(object sender, MouseEventArgs e)
        {
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                var now = DateTime.Now;
                
                // 一定間隔以内に何度も変更はできないようにする
                if ((now - lastFontResizedTime).TotalMilliseconds < 100)    // 250[ms]には1回のみ
                {
                    return;
                }

                lastFontResizedTime = now;

                int step = (e.Delta > 10 ? 1 : e.Delta < -10 ? -1 : 0);    // +1 か -1
                float currentSize = textBoxMain.Font.SizeInPoints;

                float newSize = currentSize;
                if (step < 0) {
                    // 一段階小さくする
                    foreach (var size in FontSizeArray) {
                        if (size >= currentSize) break;
                        newSize = size;
                    }
                }
                else if (step > 0)
                {
                    // 一段階大きくする
                    foreach (var size in FontSizeArray.Reverse())
                    {
                        if (size <= currentSize) break;
                        newSize = size;
                    }
                }

                if (newSize != currentSize)
                {
                    var font = textBoxMain.Font;
                    textBoxMain.Font = new System.Drawing.Font(
                        font.FontFamily, newSize, font.Style, font.Unit, font.GdiCharSet, font.GdiVerticalFont
                        );

                    if (MySettings.Font != null)
                    {
                        MySettings.Font = textBoxMain.Font;
                    }
                }
            }
        }
    }
}
