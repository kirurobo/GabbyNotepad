using SpeechLib;
using System;
using System.IO;
using System.Windows.Forms;

namespace SpeechNotepad
{
    public partial class FormMain : Form
    {
        SpeechVoiceSpeakFlags SpFlags;
        SpVoice Voice;
        ISpeechObjectTokens Tokens;

        bool IsTextChanged = false;
        string FilePath = "";

        public FormMain()
        {
            InitializeComponent();

            // SpVoiceを準備
            try
            {
                SpFlags = SpeechVoiceSpeakFlags.SVSFlagsAsync;
                Voice = new SpVoice();
            }
            catch (Exception error)
            {
                MessageBox.Show("Speech API error", error.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw;
            }

            ToolStripItemCollection collection = voiceToolStripMenuItem.DropDownItems;
            collection.Clear();

            Tokens = Voice.GetVoices("", "");
            foreach (ISpeechObjectToken token in Tokens)
            {
                ToolStripMenuItem item = new ToolStripMenuItem(token.GetDescription(System.Globalization.CultureInfo.CurrentCulture.LCID));
                collection.Add(item);
                if (token.Id == Voice.Voice.Id)
                {
                    item.Checked = true;
                }
                item.Click += new EventHandler(voiceToolStripMenuItem_Click);
            }
        }

        private void wordWrapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMain.WordWrap = wordWrapToolStripMenuItem.Checked = !wordWrapToolStripMenuItem.Checked;
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMain.SelectAll();
        }

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

                    // カーソル位置の手前にある単語を取得
                    while (curpos >= 0)
                    {
                        char chr = textBoxMain.Text[curpos];
                        if (char.IsWhiteSpace(chr)) break;
                        word = chr + word;
                        curpos--;
                    }

                    // カーソル前に単語があれば読み上げ
                    if (word.Length > 0)
                    {
                        Voice.Speak(word, SpFlags);
                    }
                }
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void textBoxMain_TextChanged(object sender, EventArgs e)
        {
            IsTextChanged = true;
        }

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

        private void fontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FontDialog fontDialog = new FontDialog();
            fontDialog.Font = textBoxMain.Font;

            if (fontDialog.ShowDialog(this) == DialogResult.OK)
            {
                textBoxMain.Font = fontDialog.Font;
            }
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMain.Cut();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMain.Copy();
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBoxMain.Paste();
        }

        private void speakParagraphToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (textBoxMain.SelectionLength == 0)
            {
                SelectSentence();
            }
            
            // カーソル前に単語があれば読み上げ
            if (textBoxMain.SelectionLength > 0)
            {
                Voice.Speak(textBoxMain.SelectedText, SpFlags);
            }
        }

        private void SelectSentence()
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

            // 文字列を選択
            textBoxMain.Select(start, end - start + 1);
        }

        private void speakAllTextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Voice.Speak(textBoxMain.Text, SpFlags);
        }

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

        private void stopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Voice.Speak("", SpeechVoiceSpeakFlags.SVSFPurgeBeforeSpeak);
        }

        private void voiceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int index = 0;
            foreach (ToolStripMenuItem item in voiceToolStripMenuItem.DropDownItems)
            {
                if (sender.Equals(item))
                {
                    item.Checked = true;
                    Voice.Voice = Tokens.Item(index);
                }
                else
                {
                    item.Checked = false;
                }
                index++;
            }
        }
    }
}
