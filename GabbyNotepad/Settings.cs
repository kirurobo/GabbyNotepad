using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.IO;
using System.Runtime.Serialization.Json;
using System.ComponentModel;

namespace GabbyNotepad
{
    [DataContract]
    public class Settings
    {
        /// <summary>
        /// 行の折り返しをするか
        /// </summary>
        [DataMember]
        public bool EnableWordWrap { get; set; } = false;

        /// <summary>
        /// 逐次発声するか
        /// </summary>
        [DataMember]
        public bool EnableWordByWord { get; set; } = true;

        /// <summary>
        /// 利用音声の名称
        /// </summary>
        [DataMember]
        public string VoiceName { get; set; } = "";

        /// <summary>
        /// フォント
        /// </summary>
        internal System.Drawing.Font Font { get; set; } = null;

        [DataMember]
        private string FontString { get; set; } = "";


        private static Encoding DataEncoding = Encoding.UTF8;
        private TypeConverter FontConverter = TypeDescriptor.GetConverter(typeof(System.Drawing.Font));


        /// <summary>
        /// 初期値に戻す
        /// </summary>
        public void Clear()
        {
            Clone(new Settings());
        }

        /// <summary>
        /// 指定ファイルから読み込み
        /// </summary>
        /// <param name="path"></param>
        public void Load(string path)
        {
            Clone(LoadFile(path));
            Font = (System.Drawing.Font)FontConverter.ConvertFromString(FontString);

            //try
            //{
            //    Font = (System.Drawing.Font)FontConverter.ConvertFromString(FontString);
            //}
            //catch(Exception e)
            //{
            //    Console.WriteLine(e.Message);
            //}
        }

        /// <summary>
        /// 指定ファイルに保存
        /// </summary>
        /// <param name="path"></param>
        public void Save(string path)
        {
            FontString = FontConverter.ConvertToString(Font);
            SaveFile(path, this);
        }

        /// <summary>
        /// 引数の値をthisに複製
        /// </summary>
        /// <param name="src"></param>
        private void Clone(Settings src)
        {
            this.EnableWordByWord = src.EnableWordByWord;
            this.EnableWordWrap = src.EnableWordWrap;
            this.Font = src.Font;
            this.FontString = src.FontString;
            this.VoiceName = src.VoiceName;
        }


        /// <summary>
        /// 設定を指定ファイルから読み込み
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static Settings LoadFile(string path)
        {
            if (!File.Exists(path))
            {
                return new Settings();
            }

            using (var reader = new StreamReader(path, DataEncoding))
            {
                using (var stream = new MemoryStream(DataEncoding.GetBytes(reader.ReadToEnd())))
                {
                    var json = new DataContractJsonSerializer(typeof(Settings));
                    return (Settings)json.ReadObject(stream);
                }
            }
        }

        /// <summary>
        /// 設定を指定ファイルに保存
        /// </summary>
        /// <param name="path"></param>
        /// <param name="settings"></param>
        public static void SaveFile(string path, Settings settings)
        {
            using (var stream = new MemoryStream())
            {
                var json = new DataContractJsonSerializer(typeof(Settings));
                json.WriteObject(stream, settings);

                using (var writer= new StreamWriter(path, false, DataEncoding))
                {
                    byte[] buff = stream.ToArray();
                    writer.WriteLine(DataEncoding.GetString(buff, 0, buff.Length));
                }
            }
        }
    }

}
