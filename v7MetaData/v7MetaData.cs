using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace v7MetaData
{
    #region Adler32
    public class Adler32
    {
        // parameters
        #region

        public const uint AdlerBase = 0xFFF1;
        public const uint AdlerStart = 0x0001;
        public const uint AdlerBuff = 0x0400;
        /// Adler-32 checksum value
        private uint m_unChecksumValue = 0;
        #endregion
        public uint ChecksumValue
        {
            get
            {
                return m_unChecksumValue;
            }
        }

        public bool MakeForBuff(byte[] bytesBuff, uint unAdlerCheckSum)
        {
            if (Object.Equals(bytesBuff, null))
            {
                m_unChecksumValue = 0;
                return false;
            }
            int nSize = bytesBuff.GetLength(0);
            if (nSize == 0)
            {
                m_unChecksumValue = 0;
                return false;
            }
            uint unSum1 = unAdlerCheckSum & 0xFFFF;
            uint unSum2 = (unAdlerCheckSum >> 16) & 0xFFFF;
            for (int i = 0; i < nSize; i++)
            {
                unSum1 = (unSum1 + bytesBuff[i]) % AdlerBase;
                unSum2 = (unSum1 + unSum2) % AdlerBase;
            }
            m_unChecksumValue = (unSum2 << 16) + unSum1;
            return true;
        }

        public bool Calc(byte[] bytesBuff)
        {
            return MakeForBuff(bytesBuff, AdlerStart);
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;
            if (this.GetType() != obj.GetType())
                return false;
            Adler32 other = (Adler32)obj;
            return (this.ChecksumValue == other.ChecksumValue);
        }

        public static bool operator ==(Adler32 objA, Adler32 objB)
        {
            if (Object.Equals(objA, null) && Object.Equals(objB, null)) return true;
            if (Object.Equals(objA, null) || Object.Equals(objB, null)) return false;
            return objA.Equals(objB);
        }

        public static bool operator !=(Adler32 objA, Adler32 objB)
        {
            return !(objA == objB);
        }

        public override int GetHashCode()
        {
            return ChecksumValue.GetHashCode();
        }

        public override string ToString()
        {
            if (ChecksumValue != 0)
                return ChecksumValue.ToString();
            return "Unknown";
        }
    }
    #endregion

    #region RC4 encryption
    public class RC4
    {
        public static byte[] MMSkey = { 0x60, 0x46, 0xD2, 0x72, 0x64, 0x25, 0x03, 0x00, 0x09, 0x89, 0x00, 0xC0, 0xDD, 0x3B, 0xE6, 0x36 };

        public static byte[] GMkey = {  0x34, 0x43, 0x33, 0x43, 0x30, 0x42, 0x46, 0x31, 0x31, 0x35, 0x46, 0x38, 0x42, 0x39, 0x35, 0x36, 
                                        0x36, 0x39, 0x46, 0x39, 0x46, 0x43, 0x34, 0x42, 0x36, 0x44, 0x33, 0x41, 0x39, 0x44, 0x36, 0x31, 
                                        0x34, 0x31};
        public byte[] key = null;
        public byte[] data;
        private bool decrypt;
        public bool MakeXOR = true;


        public RC4(byte[] data, byte[] key)
        {
            this.data = data;
            this.key = key;
        }

        public RC4(byte[] data)
        {
            this.data = data;
            this.key = MMSkey;
        }

        public RC4()
        {
        }

        public void Encode()
        {
            bool decrypt = data[0] == 0x25 || data[0] == 0x78;

            byte[] s = new byte[256];
            int i, j, t;

            for (i = 0; i < 256; i++)
                s[i] = (byte)i;

            j = 0;
            for (i = 0; i < 256; i++)
            {
                j = (j + s[i] + key[i % key.Length]) % 256;
                s[i] ^= s[j];
                s[j] ^= s[i];
                s[i] ^= s[j];
            }

            byte tt = s[0];
            i = j = 0;
            for (int x = 0; x < data.Length; x++)
            {

                i = (i + 1) % 256;
                j = (j + s[i]) % 256;

                s[i] ^= s[j];
                s[j] ^= s[i];
                s[i] ^= s[j];

                t = (s[i] + s[j]) % 256;
                data[x] ^= s[t];

                if (MakeXOR)
                {
                    data[x] ^= tt;

                    if (decrypt)
                        tt ^= (byte)(data[x] ^ s[t]);
                    else
                        tt = data[x];
                }
            }
        }
    }
    #endregion

    #region Расширение перечислений
    public static class EnumExtension
    {
        public static string GetDescription(this Enum value)
        {
            var type = value.GetType();
            var fieldInfo = type.GetField(value.ToString(CultureInfo.InvariantCulture));
            var attribs = fieldInfo.GetCustomAttributes(typeof(DescriptionAttribute), false) as DescriptionAttribute[];
            return attribs != null && attribs.Length > 0 ? attribs[0].Description : value.ToString();
        }
    }
    #endregion

    public class OleStorage
    {

        public delegate void Complete(bool ifComplete);
        public delegate void Progress(string message, int procent);

        #region Enums
        public enum StorageType
        {  //Контейнеры = каталоги
            MetaDataContainer,
            SubcontoContainer,
            SublistContainer,
            SubcontoGroupFolder,
            DocumentContainer,
            JournalContainer,
            ReportContainer,
            TypedTextContainer,
            UserDefContainer,
            PictureContainer,
            CalcJournalContainer,
            CalcVarContainer,
            AccountChartListContainer,
            AccountChartContainer,
            OperationListContainer,
            OperationContainer,
            GlobalDataContainer,
            ProvListContainer,
            TypedObjectContainer,
            WorkBookContainer,
            ModuleContainer,
            WorkPlaceType,
            RigthType,
            //Элементы
            MetaDataStream, //Описание метаданных
            MetaDataHolderContainer,
            GuidHistoryContainer,
            TagStream,
            MetaDataDescription, //Глобальник
            DialogEditor,
            TextDocument,
            [Description("Moxcel.Worksheet")]
            MoxcelWorksheet,
            UsersInterfaceType,
            SubUsersInterfaceType,
            MenuEditorType,
            ToolbarEditorType,
            PictureGalleryContainer
        }

        public static List<StorageType> ListCatalogTypes = new List<StorageType>
        { 
           StorageType.MetaDataContainer,
           //StorageType.MetaDataHolderContainer,
           StorageType.SubcontoContainer,
           StorageType.SublistContainer,
           StorageType.SubcontoGroupFolder,
           StorageType.DocumentContainer,
           StorageType.JournalContainer,
           StorageType.ReportContainer,
           StorageType.TypedTextContainer,
           StorageType.UserDefContainer,
           StorageType.PictureContainer,
           StorageType.CalcJournalContainer,
           StorageType.CalcVarContainer,
           StorageType.AccountChartListContainer,
           StorageType.AccountChartContainer,
           StorageType.OperationListContainer,
           StorageType.OperationContainer,
           StorageType.GlobalDataContainer,
           StorageType.ProvListContainer,
           StorageType.TypedObjectContainer,
           StorageType.WorkBookContainer,
           StorageType.ModuleContainer,
           StorageType.WorkPlaceType,
           StorageType.RigthType,
           StorageType.UsersInterfaceType,
           StorageType.SubUsersInterfaceType
        };

        public enum ValueType
        {
            U, // Неопределенный
            N, // Число
            S, // Строка
            D, // Дата
            E, // Перечисление
            B, // Справочник
            O, // Документ
            C, // Календарь
            A, // ВидРасчета
            T, // Счет
            K, // ВидСубконто
            P // ПланСчетов
        }
        #endregion

        public static Adler32 CRC = new Adler32();

        #region Физическая структура Метаданных

        public class MetaDescriptor
        {
            public int orderid;
            public StorageType Type;
            public MetaDescriptor Parent;
            public string Path;
            public string Name;
            public string Description;
            public string Prop_4;
            public bool isContainer;

            public MetaDescriptor()
            {
            }

            public MetaDescriptor(StorageType Type, string Name, string Description)
            {
                this.Type = Type;
                this.Name = Name;
                this.Description = Description;
                if (Type == StorageType.MoxcelWorksheet && this.Description == "Moxel WorkPlace")
                    this.Description = "Таблица";
            }

            public MetaDescriptor(StorageType Type, string Name, string Description, MetaDescriptor Parent)
                : this(Type, Name, Description)
            {
                this.Parent = Parent;
                if (Type == StorageType.MoxcelWorksheet && this.Description == "Moxel WorkPlace")
                    this.Description = "Таблица";
            }


            public static implicit operator string(MetaDescriptor Item)
            {
                return Item.ToString();
            }

            public string ToString()
            {
                return string.Format("{{\"{0}\",\"{1}\",\"{2}\",\"{3}\"}}", Type.ToString(), Name, Description, Prop_4);
            }
        }


        public enum SubType
        {
            Procedure,
            Function
        }

        public class Sub
        {
            public string PreComment = null;
            public string Name = null;
            public string Body = null;
            public SubType Type = SubType.Function;
            public List<string> Parameters = new List<string>();
            public bool Public = false;
            public bool PreDeclared = false;
            public string Tail = String.Empty;

            public List<string> Developers = new List<string>();
            public List<string> Incidents = new List<string>();
            public Dictionary<string, int> Modifacations = new Dictionary<string, int>();

            public static Dictionary<string, string> DeveloperID = new Dictionary<string, string>
            {

            };

            private int SubstringCount(string sourcestring, string substring)
            {
                return (sourcestring.Length - sourcestring.Replace(substring, "").Length) / substring.Length;
            }

            public void ParceText(string Body)
            {
                string[] splitter = { "\r\n" };
                string[] BodyStrings = Body.Split(splitter, StringSplitOptions.RemoveEmptyEntries);

                string EndOfsub;
                string TypeOfSub;

                string tempBodyString = string.Empty;

                if (!PreDeclared)
                    foreach (string BodyString in BodyStrings)
                    {
                        if (BodyString.Length < 2)
                            continue;

                        if (BodyString.Substring(0, 2) == "//")
                            continue;

                        if (BodyString.Contains("(") || tempBodyString.Contains("("))
                        {
                            if (!BodyString.Contains(")"))
                            {
                                tempBodyString += BodyString.Replace("\t", "");
                                continue;
                            }

                            tempBodyString += BodyString.Replace("\t", "");

                            string[] parameters = tempBodyString.Substring(tempBodyString.IndexOf('(') + 1, tempBodyString.IndexOf(')') - tempBodyString.IndexOf('(') - 1).Split(',');

                            foreach (string parameter in parameters)
                            {
                                if (parameter != "")
                                    Parameters.Add(parameter.Replace("Знач ", "").Split('=')[0].TrimStart(' ').TrimEnd(' '));
                            }

                            Public = tempBodyString.ToLower().Contains("экспорт");
                            PreDeclared = tempBodyString.ToLower().Contains("далее");

                            if (PreDeclared)
                            {
                                EndOfsub = "Далее";
                                string bodyToParce = string.Empty;
                                foreach (string substr in Body.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
                                {
                                    if (substr.Length > 2)
                                        if (substr.Substring(0, 2) != "//")
                                            bodyToParce += substr + "\r\n";
                                }

                                Tail = bodyToParce.Substring(bodyToParce.IndexOf(EndOfsub) + EndOfsub.Length);
                                return;
                            }

                            tempBodyString = string.Empty;
                            break;
                        }
                    }

                if (Type == SubType.Function)
                {
                    EndOfsub = "КонецФункции";
                    TypeOfSub = "Функция";
                }
                else
                {
                    EndOfsub = "КонецПроцедуры";
                    TypeOfSub = "Процедура";
                }

                if (Body.ToLower().Contains("\r\n" + EndOfsub.ToLower()))
                {

                    this.PreComment = Body.Substring(0, 2 + Body.IndexOf("\r\n" + TypeOfSub, StringComparison.OrdinalIgnoreCase));
                    this.Body = Body.Substring(PreComment.Length, Body.IndexOf("\r\n" + EndOfsub, StringComparison.OrdinalIgnoreCase) + EndOfsub.Length + 2 - PreComment.Length);
                    Tail = Body.Substring(Body.IndexOf(EndOfsub, StringComparison.OrdinalIgnoreCase) + EndOfsub.Length);

                    foreach (string ID in DeveloperID.Keys)
                    {
                        if (Body.Contains(ID))
                        {
                            Developers.Add(DeveloperID[ID]);

                            if (Modifacations.ContainsKey(DeveloperID[ID]))
                                Modifacations[DeveloperID[ID]] += SubstringCount(this.PreComment + "\r\n" + this.Body, ID);
                            else
                                Modifacations.Add(DeveloperID[ID], SubstringCount(this.PreComment + "\r\n" + this.Body, ID));
                        }
                    }

                    if (Tail.IndexOf("\r\n//+") >= 0)
                    {
                        Tail = Tail.Substring(Tail.IndexOf("\r\n//+") + 2);
                    }
                    else
                    {
                        if (Tail.IndexOf("\r\n//*") >= 0)
                        {
                            Tail = Tail.Substring(Tail.IndexOf("\r\n//*") + 2);
                        }
                        else
                        {
                            if (Tail.IndexOf("\r\n///") >= 0)
                            {
                                Tail = Tail.Substring(Tail.IndexOf("\r\n///") + 2);
                                Tail = Tail.Substring(Tail.IndexOf("\r\n") + 2);
                            }
                        }
                    }


                }
            }

            public Sub(SubType Type, string Body, string Name)
            {
                this.Type = Type;
                this.Name = Name;
                ParceText(Body);
            }

        }

        public class ProgramModule
        {
            public string GlobalVars = null;
            public string GlobalContext = null;
            public Dictionary<string, Sub> Procedures = new Dictionary<string, Sub>();
            public string Text = null;

            private void ParceProcedures(string ModuleText)
            {
                string[] splitter = { "\r\nПроцед", "\r\nФунк" };
                string[] Procedures_txt = ModuleText.Split(splitter, StringSplitOptions.RemoveEmptyEntries);

                string previous = string.Empty;
                foreach (string procedure in Procedures_txt)
                {
                    if (procedure.Length < 4)
                        continue;

                    SubType TypeOfSub = SubType.Function;
                    string EndOfsub = null;
                    string SubKeyWord = null;

                    switch (procedure.Substring(0, 4))
                    {
                        case "ция ":
                            {
                                TypeOfSub = SubType.Function;
                                EndOfsub = "КонецФункции";
                                SubKeyWord = "Функция";
                                break;
                            }
                        case "ура ":
                            {
                                TypeOfSub = SubType.Procedure;
                                EndOfsub = "КонецПроцедуры";
                                SubKeyWord = "Процедура";
                                break;
                            }
                        default:
                            {
                                if (procedure.ToLower().Contains("перем "))
                                {
                                    //Вырежем комментарии
                                    foreach (string str in procedure.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
                                    {
                                        if (str.Substring(0, 1) != "/")
                                            GlobalVars += str + "\r\n";
                                    }
                                    previous = null;
                                }
                                continue;
                            }
                    }

                    Sub Procedure;
                    string proceduretext = string.Empty;

                    proceduretext = previous + "\r\n" + SubKeyWord + " " + procedure.Substring(4);

                    string Name = SubKeyWord + " " + procedure.Substring(4, procedure.IndexOf('(') - 4).TrimStart();

                    if (Procedures.ContainsKey(Name))
                    {
                        Procedure = Procedures[Name];
                        Procedure.ParceText(proceduretext);
                    }
                    else
                    {
                        Procedure = new Sub(TypeOfSub, proceduretext, Name);
                        Procedures.Add(Procedure.Name, Procedure);
                    }

                    previous = Procedure.Tail;

                    if (previous.ToLower().Contains("перем "))
                    {
                        GlobalVars += previous;
                        previous = null;
                    }
                }

                if (previous != null)
                    foreach (string str in previous.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
                    {   //Вырежем комментарии
                        if (str.Substring(0, 1) != "/")
                            GlobalContext += str + "\r\n";
                    }
            }

            public ProgramModule(string text)
            {
                this.Text = text;
                ParceProcedures(text);
            }
        }

        public class MetaItem : MetaDescriptor
        {

            private byte[] Decompress(byte[] compressed)
            {
                MemoryStream ms = new MemoryStream(compressed.Length);
                ms.Write(compressed, 0, compressed.Length);
                ms.Position = 0;
                DeflateStream compressedzipStream = new DeflateStream(ms, System.IO.Compression.CompressionMode.Decompress, false);

                byte[] data = new byte[compressed.Length * 10];
                int totalCount = 0;
                int bytesRead = compressedzipStream.Read(data, 0, data.Length);
                totalCount += bytesRead;
                compressedzipStream.Close();
               //compressedzipStream.Dispose();
                Array.Resize<byte>(ref data, totalCount);
                return data;
            }

            static List<StorageType> Modules = new List<StorageType> 
            { 
                StorageType.TextDocument, 
                StorageType.MetaDataDescription
            };

            static List<StorageType> Dialogs = new List<StorageType> 
            { 
                StorageType.DialogEditor
            };

            static List<StorageType> Moxels = new List<StorageType> 
            { 
                StorageType.MoxcelWorksheet
            };

            public uint CheckSumm = 0;
            bool IsCompressed = false;
            bool IsEncrypted = false;
            public string Moduletext = string.Empty;
            public ProgramModule Module;
            private byte[] _data;
            public byte[] data
            {
                get { return _data; }

                set
                {

                    CheckSumm = 0;
                    if (IsCompressed)
                    {
                        byte[] decompressed = Decompress(value);

                        if (decompressed.Length > 0)
                        {
                            if (Type == StorageType.MetaDataDescription)
                            {
                                IsEncrypted = decompressed[0] == 0x9e; // определяем зашифрован или нет

                                if (IsEncrypted) 
                                {
                                    byte[] Crypted = new byte[decompressed.Length];
                                   
                                    //У шифрованного глобальника первые 510 байт - хз что, после обрезки расшифровывается успешно.
                                    Array.ConstrainedCopy(decompressed, 510, Crypted, 0, decompressed.Length - 510);
                                    decompressed = Crypted;

                                    RC4 Codec = new RC4(decompressed, RC4.GMkey);
                                    Codec.MakeXOR = false;      //Тут используется не модифицированый RC4
                                    Codec.Encode();
                                }

                            }
                        }

                        Moduletext = Encoding.GetEncoding(1251).GetString(decompressed);
                        Module = new ProgramModule(Moduletext);
                    }

                    _data = value;

                    if (Type == StorageType.MetaDataStream)
                    {
                        IsEncrypted = _data[0] != 0xFF;
                        if (IsEncrypted)
                        {
                            RC4 Codec = new RC4(_data, RC4.MMSkey);
                            Codec.Encode();
                        }

                        Moduletext = Encoding.GetEncoding(1251).GetString(_data);
                    }

                    if (Type == StorageType.TagStream)
                    {
                        IsEncrypted = true; //Это всегда зашифровано.
                        RC4 Codec = new RC4(_data, RC4.MMSkey);
                        Codec.Encode();

                        Moduletext = Encoding.GetEncoding(1251).GetString(_data);
                    }

                    if (Type == StorageType.DialogEditor)
                    {
                        Moduletext = Encoding.GetEncoding(1251).GetString(_data); //Надо распарсить в XML
                    }


                    if (Moduletext != null && Moduletext != string.Empty)
                    {
                        CheckSumm = (uint)Moduletext.GetHashCode();
                    }

                    if (CheckSumm == 0)
                    {
                        CRC.Calc(value);
                        CheckSumm = CRC.ChecksumValue;
                    }

                    _data = value;

                }
            }


            public MetaItem(StorageType Type, string Name, string Description)
                : base(Type, Name, Description)
            {
                IsCompressed = Modules.Contains(Type);
                isContainer = false;
            }

            public MetaItem(StorageType Type, string Name, string Description, MetaDescriptor Parent)
                : base(Type, Name, Description, Parent)
            {
                IsCompressed = Modules.Contains(Type);
                isContainer = false;
            }

        }

        public class MetaContainer : MetaDescriptor
        {
            public List<MetaDescriptor> Items = null;
            public string Contents
            {
                get
                {
                    string res = "{Container.Contents";
                    foreach (MetaDescriptor Item in Items)
                    {
                        res += "," + Item.ToString();
                    }
                    res += "}";
                    return res;
                }
            }

            public MetaItem GetSubItem(string path)
            {
                string[] PathDirectories = path.Split('\\');

                MetaContainer Root = this;

                foreach (string subdir in PathDirectories)
                {
                    if (subdir != "")
                    {
                        if (Name == subdir)
                            continue;
                        MetaDescriptor item = Root.Items.Find(x => x.Name == subdir);
                        if (item == null)
                            return null;
                        if (item.isContainer)
                            Root = (MetaContainer)item;
                        else
                        {
                            return (MetaItem)item;
                        }


                    }
                }
                return null;
            }

            public MetaContainer GetSubCatalog(string path)
            {
                string[] PathDirectories = path.Split('\\');

                MetaContainer Root = this;

                foreach (string subdir in PathDirectories)
                {
                    if (subdir != "")
                    {
                        if (Name == subdir)
                            continue;
                        MetaDescriptor item = Root.Items.Find(x => x.Name == subdir);
                        if (item == null)
                            return null;

                        if (item.isContainer)
                            Root = (MetaContainer)item;
                        else
                        {
                            return (MetaContainer)item;
                        }


                    }
                }
                return Root;
            }

            public MetaContainer(StorageType Type, string Name, string Description)
                : base(Type, Name, Description)
            {
                Items = new List<MetaDescriptor>();
                isContainer = true;
            }

            public MetaContainer(StorageType Type, string Name, string Description, MetaDescriptor Parent)
                : base(Type, Name, Description, Parent)
            {
                Items = new List<MetaDescriptor>();
                isContainer = true;
            }


            public List<MetaDescriptor> ParceContainer_Contents(string Contents, MetaDescriptor Prent)
            {
                List<MetaDescriptor> Items = new List<MetaDescriptor>();
                Contents = Contents.Replace("{\"Container.Contents\",", "");
                Contents = Contents.Substring(0, Contents.Length - 2);
                StorageType Type;
                foreach (string subitem in Contents.Replace("},{", "#").Split('#'))
                {
                    string[] subelements = subitem.Replace("{", "").Replace("}", "").Replace("\"", "").Split(',');
                    if (StorageType.TryParse(subelements[0].Replace(".", ""), out Type))
                    {
                        if (ListCatalogTypes.Contains(Type))
                            Items.Add(new MetaContainer(Type, subelements[1], subelements[2], Parent));
                        else
                            Items.Add(new MetaItem(Type, subelements[1], subelements[2], Parent));
                    }

                }


                return Items;
            }

            bool ReadStorage(IStorage RootStorage)
            {
                IStorage storage = null;
                IStream pIStream = null;
                IEnumSTATSTG pIEnumStatStg = null;
                byte[] data = { 0 };
                uint fetched = 0;

                System.Runtime.InteropServices.ComTypes.STATSTG[] regelt =
                {
                    new System.Runtime.InteropServices.ComTypes.STATSTG()
                };

                RootStorage.OpenStream("Container.Contents", IntPtr.Zero, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE), 0, out pIStream);

                data = ReadIStream(pIStream);
                Marshal.ReleaseComObject(pIStream);
                Marshal.FinalReleaseComObject(pIStream);
                pIStream = null;

                Items = ParceContainer_Contents(Encoding.GetEncoding(1251).GetString(data), this);

                RootStorage.EnumElements(0, IntPtr.Zero, 0, out pIEnumStatStg);
                while (pIEnumStatStg.Next(1, regelt, out fetched) == 0)
                {
                    string filePage = regelt[0].pwcsName;
                    if (filePage != "Container.Contents")
                    {

                        if ((STGTY)regelt[0].type == STGTY.STGTY_STREAM)
                        {
                            MetaItem item = (MetaItem)Items.Find(x => x.Name == filePage);
                            if (item != null)
                            {
                                RootStorage.OpenStream(filePage, IntPtr.Zero, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE),
                                    0,
                                    out pIStream);
                                if (pIStream != null)
                                {
                                    item.Path += Path + "\\" + Name;
                                    item.data = ReadIStream(pIStream);
                                    Marshal.ReleaseComObject(pIStream);
                                    Marshal.FinalReleaseComObject(pIStream);
                                    pIStream = null;
                                }
                            }
                            else
                            {
                                //if (filePage != "MetaContainer.Profile")
                                //    MessageBox.Show("В структуре не найден объект " + filePage);
                            }

                        }

                        if ((STGTY)regelt[0].type == STGTY.STGTY_STORAGE)
                        {
                            MetaContainer item = (MetaContainer)Items.Find(x => (x.Name == filePage) && (x.isContainer));
                            if (item != null)
                            {
                                RootStorage.OpenStorage(filePage, null, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE),
                                    IntPtr.Zero, 0, out storage);
                                item.Path += Path + "\\" + Name;
                                item.Parent = this;
                                item.ReadStorage(storage);
                                Marshal.ReleaseComObject(storage);
                                Marshal.FinalReleaseComObject(storage);
                                storage = null;
                            }
                            else
                            {
                                // MessageBox.Show("В структуре не найден объект " + filePage);
                            }
                        }
                    }
                }
                return true;
            }

            public MetaContainer(string mdFileNAme, Progress progressproc = null)
            {
                IStorage storage = null;
                IStorage RootStorage = null;
                IStream pIStream = null;
                IEnumSTATSTG pIEnumStatStg = null;
                byte[] data = { 0 };
                uint fetched = 0;

                if (NativeMethods.StgOpenStorage(mdFileNAme, null, STGM.READ | STGM.SHARE_DENY_WRITE, IntPtr.Zero, 0, out RootStorage) == 0)
                {

                    System.Runtime.InteropServices.ComTypes.STATSTG[] regelt =
                    {
                        new System.Runtime.InteropServices.ComTypes.STATSTG()
                    };

                    Name = "Root";

                    RootStorage.OpenStream("Container.Contents", IntPtr.Zero, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE), 0, out pIStream);
                    data = ReadIStream(pIStream);
                    Marshal.ReleaseComObject(pIStream);
                    Marshal.FinalReleaseComObject(pIStream);
                    pIStream = null;

                    Items = ParceContainer_Contents(Encoding.GetEncoding(1251).GetString(data), this);
                    RootStorage.EnumElements(0, IntPtr.Zero, 0, out pIEnumStatStg);
                    while (pIEnumStatStg.Next(1, regelt, out fetched) == 0)
                    {
                        string filePage = regelt[0].pwcsName;
                        if (filePage != "Container.Contents")
                        {
                            if ((STGTY)regelt[0].type == STGTY.STGTY_STORAGE)
                            {
                                MetaContainer item = (MetaContainer)Items.Find(x => x.Name == filePage);

                                if (item != null)
                                {
                                    item.Path += Path + "\\" + Name;
                                    RootStorage.OpenStorage(filePage, null, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE),
                                        IntPtr.Zero, 0, out storage);
                                    item.Parent = this;
                                    item.ReadStorage(storage);
                                    Marshal.ReleaseComObject(storage);
                                    Marshal.FinalReleaseComObject(storage);
                                    storage = null;
                                }
                                else
                                {
                                    //throw new Exception("В структуре не найден объект " + filePage);
                                }

                            }

                            if ((STGTY)regelt[0].type == STGTY.STGTY_STREAM)
                            {
                                MetaItem item = (MetaItem)Items.Find(x => x.Name == filePage);
                                if (item != null)
                                {
                                    RootStorage.OpenStream(filePage, IntPtr.Zero, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE),
                                        0,
                                        out pIStream);
                                    if (pIStream != null)
                                    {
                                        item.Path += Path + "\\" + Name;
                                        item.data = ReadIStream(pIStream);
                                        Marshal.ReleaseComObject(pIStream);
                                        Marshal.FinalReleaseComObject(pIStream);
                                        pIStream = null;
                                    }
                                }
                                else
                                {
                                    //if (filePage != "MetaContainer.Profile")
                                    //    MessageBox.Show("В структуре не найден объект " + filePage);
                                }

                            }

                        }
                    }

                    if (pIStream != null)
                    {
                        Marshal.ReleaseComObject(pIStream);
                        Marshal.FinalReleaseComObject(pIStream);
                        pIStream = null;
                    }

                    if (storage != null)
                    {
                        Marshal.ReleaseComObject(storage);
                        Marshal.FinalReleaseComObject(storage);
                        storage = null;
                    }

                    if (RootStorage != null)
                    {
                        Marshal.ReleaseComObject(RootStorage);
                        Marshal.FinalReleaseComObject(RootStorage);
                        RootStorage = null;
                    }
                }

                else
                {
                    throw new Exception("Не удалось открыть конфигурацию");
                }
            }

        }

        #endregion

        #region Логическая структура Метаданных
        public class MetaObject
        {
            public int ID = 0;
            public string Identity = string.Empty;
            public string Alias = string.Empty;
            public string Description = string.Empty;

            public override string ToString()
            {
                return string.Format("{{\"{0}\",\"{1}\",\"{2}\",\"{3}\"}}", ID, Identity, Alias, Description);
            }

            public static implicit operator string(MetaObject obj)
            {
                return obj.ToString();
            }

        }

        public class Form : MetaObject
        {
            MetaContainer Catalog;
            public MetaItem Dialog;
            public MetaItem DialogModule;
            public List<MetaItem> Moxels = new List<MetaItem>();

            public Form()
                : base()
            {
            }

            public Form(MetaContainer Catalog)
            {
                this.Catalog = Catalog;
                this.DialogModule = (MetaItem)Catalog.Items.Find(x => x.Type == StorageType.TextDocument); //.GetSubItem("\\MD Programm text");
                this.Dialog = (MetaItem)Catalog.Items.Find(x => x.Type == StorageType.DialogEditor);// Catalog.GetSubItem("\\Dialog Stream");
                foreach (MetaItem item in Catalog.Items.FindAll(x => x.Type == StorageType.MoxcelWorksheet))
                    Moxels.Add(item);

            }

        }

        public class Parameter : MetaObject
        {
            public ValueType Type;
            public int Length = 0;
            public int Precision = 0;
            public MetaObject TypedObject;
            public bool NoNegative = false;
            public bool Splittriades = false;

            public override string ToString()
            {
                return string.Format("{{\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{7}\",\"{4}\",\"{5}\",\"{6}\"}}", ID, Identity, Alias, Description, Length, Precision, TypedObject.ID, Type.ToString());
            }

        }

        public enum NumeratorType
        {
            Text = 1,
            Number = 2
        }

        public enum AutoNumeration
        {
            No = 1,
            Yes = 2
        }

        public class SubcontoParameter : Parameter
        {
            public bool Perodical = false;
            public bool UseForItem = false;
            public bool UseForGroup = false;
            public bool Sort = false;
            public bool HistoryManual = false;
            public bool ChangeByDocument = false;
            public bool Selection = false;

        }

        public class Subconto : MetaObject
        {
            public int Parent = 0;
            public int CodeLength = 0;
            public int CodeSeries = 1;
            public NumeratorType CodeType = NumeratorType.Text;
            public AutoNumeration AutoNum = AutoNumeration.Yes;
            public int NameLength = 0;
            public int MainRepresent = 0;
            public int EditMode = 0;
            public int LevelCount = 0;
            public int SelectFormID = 0;
            public int MainFormID = 0;
            public int OneForm = 0;
            public int UniqCode = 0;
            public int GroupsInTop = 0;

            public Form SelectForm = null;
            public Form MainForm = null;
            public Form Form;
            public Form FolderForm;
            public List<Form> ListForms = new List<Form>();
            public List<SubcontoParameter> Params = new List<SubcontoParameter>();

            public Subconto(string[] objparams) //Справочники
            {
                ID = int.Parse(objparams[0]);
                Identity = objparams[1];
                Description = objparams[2];
                Alias = objparams[3];
                Parent = int.Parse(objparams[4]);
                CodeLength = int.Parse(objparams[5]);
                CodeSeries = int.Parse(objparams[6]);
                CodeType = (NumeratorType)int.Parse(objparams[7]);
                AutoNum = (AutoNumeration)int.Parse(objparams[8]);
                NameLength = int.Parse(objparams[9]);
                MainRepresent = int.Parse(objparams[10]);
                EditMode = int.Parse(objparams[11]);
                LevelCount = int.Parse(objparams[12]);
                SelectFormID = int.Parse(objparams[13]);
                MainFormID = int.Parse(objparams[14]);
                OneForm = int.Parse(objparams[15]);
                UniqCode = int.Parse(objparams[16]);
                GroupsInTop = int.Parse(objparams[17].Replace(",", ""));
            }
        }

        public class DocumentParameter : Parameter
        {
        }

        public class DocumentTableParameter : Parameter
        {
            public bool TotalColumn = false;
        }

        public class Document : MetaObject
        {
            public enum DocumentPeriodicity
            {
                All = 0,
                InYead = 1,
                InQuartal,
                InMonth,
                InDay
            }

            public int NumberLength = 0;
            public DocumentPeriodicity Periodicity = 0;    //Периодичность: 0 - по всем данного вида, 1 - в пределах года, 2 - в пределах квартала, 3 - в пределах месяца, 4 - в пределах дня.
            public NumeratorType NumberType = NumeratorType.Text;
            public AutoNumeration AutoNum = AutoNumeration.Yes;
            public int JournalID = 0;
            public int Unkown = -1;
            public bool UniqNumber = true;
            public int NumeratorID = 0;

            public bool Operative = false;
            public bool Calculation = false;
            public bool Accounting = false;
            public List<Document> InputOnBasis;
            public bool BaseForAnyDocument = false;
            public int CreateOperation = 2;
            public bool AutoLineNumber = true;
            public bool AutoRemovActions = true;
            public bool EditOperation = false;
            public bool CanDoActions = true;


            public Form Form;
            public MetaItem TransactionModule;
            public List<DocumentParameter> HeadFileds = new List<DocumentParameter>();
            public List<DocumentTableParameter> TableFields = new List<DocumentTableParameter>();
        }

        public class Journal : MetaObject
        {
            public int Unknown = 0;
            public int JornalType = 0;//       Тип журнала: 0 - обычный, 1 - общий
            public Form SelectForm;//       Числовой идентификатор формы для выбора
            public Form MainForm;//         Числовой идентификатор основной формы  
            public bool NoAdditional;//     Не дополнительный: 0 - журнал дополнительный, 1 - нет

            public List<Form> ListForms = new List<Form>();
        }

        public class EnumVal : MetaObject
        {
            int OrderID = 1;
        }

        public class EnumItem : MetaObject
        {
            public List<EnumVal> Values = new List<EnumVal>();
        }

        public class CalculationAlgorithm : MetaObject
        {
            public MetaItem CalculationModule;
        }

        public class CalcJournal : MetaObject
        {
            public int Unknown = 0;
            public int JornalType = 0;//       Тип журнала: 0 - обычный, 1 - общий
            public Form SelectForm;//       Числовой идентификатор формы для выбора
            public Form MainForm;//         Числовой идентификатор основной формы  
            public List<Form> ListForms = new List<Form>();
        }

        public class BuhParameters : MetaObject         //Параметры бухгалтерии
        {
            public List<Form> AccountChart = new List<Form>();
            public List<Form> AccountChartList = new List<Form>();
            public List<Form> OperationList = new List<Form>();
            public List<Form> ProvListList = new List<Form>();
        }

        public class ReportItem : MetaObject
        {
            public Form Form;
        }

        public class CalcVarItem : MetaObject
        {
            public Form Form;
        }

        public enum DescriptorType
        {
            Null,
            MainDataContDef,
            TaskItem,
            GenJrnlFldDef,
            DocSelRefObj,
            DocNumDef,
            Registers,
            Documents,
            Journalisters,
            EnumList,
            ReportList,
            CJ,
            Calendars,
            Algorithms,
            RecalcRules,
            CalcVars,
            Groups,
            [Description("Document Streams")]
            DocumentStreams,
            Buh,
            CRC,
            //***************************************
            Refers, //Ссылки. Для разных типов разные
            Consts, //Константы
            SbCnts, //Справочники
            Params, //Атрибуты
            Form    //Описатель списка форм
        }

        public class ExtReport : CalcVarItem
        {
            public string FilePath = string.Empty;
            MetaContainer MetadataFileTree = null;
            string MMS = string.Empty;

            public ExtReport()
                : base()
            {

            }

            public ExtReport(string FileName)
                : base()
            {
                try
                {
                    this.FilePath = Path.GetDirectoryName(FileName);

                    Identity = Path.GetFileName(FileName);
                    Alias = Path.GetFileNameWithoutExtension(FileName);

                    MetadataFileTree = new MetaContainer(FileName);
                    if (MetadataFileTree.Items == null)
                        return;

                    MMS = MetadataFileTree.GetSubItem("Root\\Main MetaData Stream").Moduletext;
                    Form = new Form(MetadataFileTree.GetSubCatalog("Root\\"));
                }
                catch (Exception ex)
                {
                    throw new Exception("Ошибка открытия внешней обработки: " + ex.ToString());
                }
            }
        }

        public class TaskItem
        {
            public List<Document> Documents = new List<Document>();
            public List<Subconto> Subcontos = new List<Subconto>();
            public List<Journal> Journals = new List<Journal>();
            public List<ReportItem> Reports = new List<ReportItem>();
            public List<CalcVarItem> CalcVars = new List<CalcVarItem>();
            public List<CalculationAlgorithm> Algorithms = new List<CalculationAlgorithm>();
            public List<CalcJournal> CJ = new List<CalcJournal>();
            public BuhParameters Buh = new BuhParameters();
            public MetaItem GlobalModule = null;
            public MetaItem TagStream = null;
            public List<Guid> GUIDData = new List<Guid>();
            MetaContainer MetadataFileTree = null;
            public bool Isencrypted = false;
            public string MDFileName;
            public string MMS = null;

            private string CompareModules(MetaItem First, MetaItem Second)
            {
                return CompareModules(First.Module, Second.Module);
            }

            private string CompareModules(ProgramModule First, ProgramModule Second, bool GetEditors = false)
            {
                string report = string.Empty;

                Dictionary<string, Sub> Procedures = First.Procedures;
                Dictionary<string, Sub> Procedures_second = Second.Procedures;

                if (First.GlobalVars != null)
                {
                    if (Second.GlobalVars != null)
                    {
                        if (First.GlobalVars.GetHashCode() != Second.GlobalVars.GetHashCode())
                            report += "   Блок глобальных переменных модуля\n";
                    }
                    else
                    {
                        report += "   Добавлен блок глобальных переменных модуля\n";
                    }
                }
                else
                {
                    if (Second.GlobalVars != null)
                    {
                        report += "   Удален блок глобальных переменных модуля\n";
                    }
                }

                if (First.GlobalContext != null)
                {
                    if (Second.GlobalContext != null)
                    {
                        if (First.GlobalContext.GetHashCode() != Second.GlobalContext.GetHashCode())
                            report += "   Блок кода вне процедур\n";
                    }
                    else
                    {
                        report += "   Добавлен блок кода вне процедур\n";
                    }
                }
                else
                {
                    if (Second.GlobalContext != null)
                    {
                        report += "   Удален блок кода вне процедур\n";
                    }
                }



                foreach (KeyValuePair<string, Sub> Procedure in Procedures)
                {
                    if (Procedures_second.ContainsKey(Procedure.Key))
                    {
                        if (Procedure.Value.Body.GetHashCode() != Procedures_second[Procedure.Key].Body.GetHashCode())
                        {
                            report += "    " + Procedure.Key + "()\n";
                            foreach (string Parameter in Procedure.Value.Parameters)
                            {
                                if (!Procedures_second[Procedure.Key].Parameters.Contains(Parameter))
                                {
                                    report += "       Добавлен параметр \"" + Parameter + "\"\n";
                                }
                            }

                            foreach (string Parameter in Procedures_second[Procedure.Key].Parameters)
                            {
                                if (!Procedure.Value.Parameters.Contains(Parameter))
                                {
                                    report += "       Удален параметр \"" + Parameter + "\"\n";
                                }
                            }

                            string editors = null;

                            if (GetEditors)
                                foreach (KeyValuePair<string, int> mod in Procedure.Value.Modifacations)
                                {

                                    if (Procedures_second[Procedure.Key].Modifacations.ContainsKey(mod.Key))
                                    {
                                        //Если количество изменений этого разработчика больше чем в предыдущей версии - значит он автор изменений
                                        if (mod.Value > Procedures_second[Procedure.Key].Modifacations[mod.Key])
                                            editors += mod.Key + ',';
                                    }
                                    else
                                    {
                                        editors += mod.Key + ',';
                                    }
                                }

                            if (GetEditors)
                                foreach (KeyValuePair<string, int> mod in Procedures_second[Procedure.Key].Modifacations)
                                {
                                    if (!Procedure.Value.Modifacations.ContainsKey(mod.Key))
                                    {
                                        editors += mod.Key + ',';
                                    }
                                }

                            if (editors != null)
                                editors = "    " + "    " + "Авторы изменений: " + editors.Substring(0, editors.Length - 1) + "\n";

                            report += editors;
                        }
                    }
                    else
                    {
                        report += "    Добавлена " + Procedure.Key + "()\n";
                        string editors = null;

                        if (GetEditors)
                            foreach (KeyValuePair<string, int> mod in Procedure.Value.Modifacations)
                            {
                                editors += mod.Key + ',';
                            }

                        if (editors != null)
                            editors = "    " + "    " + "Авторы изменений: " + editors.Substring(0, editors.Length - 1) + "\n";

                        report += editors;
                    }
                }

                foreach (KeyValuePair<string, Sub> Procedure in Procedures_second)
                {
                    if (!Procedures.ContainsKey(Procedure.Key))
                    {
                        report += "    Удалена " + Procedure.Key + "()\n";
                    }
                }
                return report;
            }

            public string CompareWith(TaskItem Second, bool comparemodeules, bool GetEditors = false)
            {
                string report = string.Empty;

                //Сравним структуру метаданных
                if (MMS.GetHashCode() != Second.MMS.GetHashCode())
                    report += "Изменена структура метаданных\n";

                //Сравним глобальники
                #region Глобальный модуль
                if (GlobalModule.CheckSumm != Second.GlobalModule.CheckSumm)
                {
                    report += "Глобальный модуль\n";

                    if (comparemodeules)
                        report += CompareModules(GlobalModule.Module, Second.GlobalModule.Module, GetEditors);

                }
                #endregion
                
                //Сравним справочники
                #region  справочники
                foreach (var item in Subcontos)
                {
                    var item2 = Second.Subcontos.Find(x => x.Identity == item.Identity);

                    if (item2 != null)
                    {
                        if (item.Form != null)
                        {
                            if (item2.Form.DialogModule.CheckSumm != item.Form.DialogModule.CheckSumm)
                            {
                                report += string.Format("Справочник.{0}.ФормаЭлемента.Модуль\n", item.Identity);

                                if (comparemodeules)
                                    report += CompareModules(item.Form.DialogModule.Module, item2.Form.DialogModule.Module, GetEditors);

                                if (item2.Form.Dialog.CheckSumm != item.Form.Dialog.CheckSumm)
                                    report += string.Format("Справочник.{0}.ФормаЭлемента.Диалог\n", item.Identity);
                            }

                            if (item.Form.Moxels != null)
                                foreach (MetaItem Moxel in item.Form.Moxels)
                                {
                                    MetaItem Moxel2 = item2.Form.Moxels.Find(x => x.Name == Moxel.Name);
                                    if (Moxel2 != null)
                                    {
                                        if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                            report += string.Format("Таблица Справочник.{0}.ФормаЭлемента.{1}\n", item.Identity, Moxel.Description);
                                    }
                                    else
                                    {
                                        report += string.Format("Добавлена таблица Справочник.{0}.ФормаЭлемента{1}\n", item.Identity, Moxel.Description);
                                    }

                                }
                        }

                        if (item.FolderForm != null)
                        {
                            if (item2.FolderForm.DialogModule.CheckSumm != item.FolderForm.DialogModule.CheckSumm)
                            {
                                report += string.Format("Справочник.{0}.ФормаГруппы.Модуль\n", item.Identity);

                                if (comparemodeules)
                                    report += CompareModules(item.FolderForm.DialogModule.Module, item2.FolderForm.DialogModule.Module, GetEditors);

                                if (item2.FolderForm.Dialog.CheckSumm != item.FolderForm.Dialog.CheckSumm)
                                    report += string.Format("Справочник.{0}.ФормаГруппы.Диалог\n", item.Identity);
                            }

                            if (item.FolderForm.Moxels != null)
                                foreach (MetaItem Moxel in item.FolderForm.Moxels)
                                {
                                    MetaItem Moxel2 = item2.FolderForm.Moxels.Find(x => x.Name == Moxel.Name);
                                    if (Moxel2 != null)
                                    {
                                        if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                            report += string.Format("Таблица Справочник.{0}.ФормаГруппы.{1}\n", item.Identity, Moxel.Description);
                                    }
                                    else
                                    {
                                        report += string.Format("Добавлена таблица Справочник.{0}.ФормаГруппы{1}\n", item.Identity, Moxel.Description);
                                    }

                                }
                        }

                        foreach (var frm in item.ListForms)
                        {
                            var frm2 = item2.ListForms.Find(x => x.Identity == frm.Identity);
                            if (frm2 != null)
                            {
                                if (frm.DialogModule != null)
                                {
                                    if (frm.DialogModule.CheckSumm != frm2.DialogModule.CheckSumm)
                                    {
                                        report += string.Format("Справочник.{1}.ФормаСписка.{0}.Модуль\n", frm.Identity, item.Identity);
                                        if (comparemodeules)
                                            report += CompareModules(frm.DialogModule.Module, frm2.DialogModule.Module, GetEditors);

                                        if (frm.Dialog.CheckSumm != frm2.Dialog.CheckSumm)
                                            report += string.Format("Справочник.{1}.ФормаСписка.{0}.Диалог\n", frm.Identity, item.Identity);
                                    }
                                }

                                if (frm.Moxels != null)
                                    foreach (MetaItem Moxel in frm.Moxels)
                                    {
                                        MetaItem Moxel2 = frm2.Moxels.Find(x => x.Name == Moxel.Name);

                                        if (Moxel2 != null)
                                        {
                                            if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                                report += string.Format("Таблица Справочник.{1}.ФормаСписка.{0}.{2}\n", frm.Identity, item.Identity, Moxel.Description);
                                        }
                                        else
                                        {
                                            report += string.Format("Добавлена таблица Справочник.{1}.ФормаСписка.{0}.{2}\n", frm.Identity, item.Identity, Moxel.Description);
                                        }

                                    }
                            }
                            else
                                report += string.Format("Добавлена форма: Справочник.{1}.ФормаСписка.{0}\n", frm.Identity, item.Identity);
                        }
                    }
                    else
                        report += string.Format("Добавлен: Справочник.{0}\n", item.Identity);
                }
                #endregion

                //Сравним документы
                #region Документы
                foreach (var item in Documents)
                {
                    var item2 = Second.Documents.Find(x => x.Identity == item.Identity);

                    if (item2 != null)
                    {
                        if (item.Form != null)
                        {
                            if (item2.Form.DialogModule.CheckSumm != item.Form.DialogModule.CheckSumm)
                            {
                                report += string.Format("Документ.{0}.Форма.Модуль\n", item.Identity);

                                if (comparemodeules)
                                    report += CompareModules(item.Form.DialogModule.Module, item2.Form.DialogModule.Module, GetEditors);

                                if (item2.Form.Dialog.CheckSumm != item.Form.Dialog.CheckSumm)
                                    report += string.Format("Документ.{0}.Форма.Диалог\n", item.Identity);
                            }

                            if (item.Form.Moxels != null)
                                foreach (MetaItem Moxel in item.Form.Moxels)
                                {
                                    MetaItem Moxel2 = item2.Form.Moxels.Find(x => x.Name == Moxel.Name);
                                    if (Moxel2 != null)
                                    {
                                        if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                            report += string.Format("Таблица Документ.{0}.Форма.{1}\n", item.Identity, Moxel.Description);
                                    }
                                    else
                                    {
                                        report += string.Format("Добавлена таблица Документ.{0}.Форма{1}\n", item.Identity, Moxel.Description);
                                    }

                                }
                        }

                        if (item.TransactionModule != null)
                        {
                            if (item2.TransactionModule.CheckSumm != item.TransactionModule.CheckSumm)
                            {
                                report += string.Format("Документ.{0}.МодульПроведения\n", item.Identity);
                                if (comparemodeules)
                                    report += CompareModules(item.TransactionModule.Module, item2.TransactionModule.Module, GetEditors);
                            }
                        }
                    }
                    else
                        report += string.Format("Добавлен: Документ.{0}\n", item.Identity);
                }
                #endregion

                //Сравним Журналы
                #region Журналы
                foreach (var item in Journals)
                {
                    var item2 = Second.Journals.Find(x => x.Identity == item.Identity);

                    if (item2 != null)
                    {
                        foreach (var frm in item.ListForms)
                        {
                            var frm2 = item2.ListForms.Find(x => x.Identity == frm.Identity);
                            if (frm2 != null)
                            {
                                if (frm.DialogModule.CheckSumm != frm2.DialogModule.CheckSumm)
                                {
                                    report += string.Format("Журнал.{0}.ФормаСписка.{1}.Модуль\n", frm.Identity, item.Identity);

                                    if (comparemodeules)
                                        report += CompareModules(frm.DialogModule.Module, frm2.DialogModule.Module, GetEditors);

                                    if (frm.Dialog.CheckSumm != frm2.Dialog.CheckSumm)
                                        report += string.Format("Журнал.{0}.ФормаСписка.{1}.Диалог\n", frm.Identity, item.Identity);
                                }

                                if (frm.Moxels != null)
                                    foreach (MetaItem Moxel in frm.Moxels)
                                    {
                                        MetaItem Moxel2 = frm2.Moxels.Find(x => x.Name == Moxel.Name);

                                        if (Moxel2 != null)
                                        {
                                            if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                                report += string.Format("Таблица Журнал.{1}.ФормаСписка.{0}.{2}\n", frm.Identity, item.Identity, Moxel.Description);
                                        }
                                        else
                                        {
                                            report += string.Format("Добавлена таблица Журнал.{1}.ФормаСписка.{0}.{2}\n", frm.Identity, item.Identity, Moxel.Description);
                                        }

                                    }

                            }
                            else
                                report += string.Format("Добавлена форма: Журнал.{0}.ФормаСписка.{1}\n", frm.Identity, item.Identity);
                        }
                    }
                    else
                        report += string.Format("Добавлен: Журнал.{0}\n", item.Identity);
                }
                #endregion

                //Сравним виды расчетов
                #region Отчеты
                foreach (var item in Algorithms)
                {
                    var item2 = Second.Algorithms.Find(x => x.Identity == item.Identity);

                    if (item2 != null)
                    {
                        if (item2.CalculationModule.CheckSumm != item.CalculationModule.CheckSumm)
                        {
                            report += string.Format("ВидРасчета.{0}.МодульРасчета\n", item.Identity);
                            if (comparemodeules)
                                report += CompareModules(item.CalculationModule.Module, item2.CalculationModule.Module, GetEditors);
                        }
                    }
                    else
                        report += string.Format("Добавлен: ВидРасчета.{0}\n", item.Identity);
                }
                #endregion               

                //Сравним Журналы Расчетов
                #region Журналы
                foreach (var item in CJ)
                {
                    var item2 = Second.Journals.Find(x => x.Identity == item.Identity);

                    if (item2 != null)
                    {
                        foreach (var frm in item.ListForms)
                        {
                            var frm2 = item2.ListForms.Find(x => x.Identity == frm.Identity);
                            if (frm2 != null)
                            {
                                if (frm.DialogModule.CheckSumm != frm2.DialogModule.CheckSumm)
                                {
                                    report += string.Format("ЖурналРасчетов.{0}.ФормаСписка.{1}.Модуль\n", frm.Identity, item.Identity);

                                    if (comparemodeules)
                                        report += CompareModules(frm.DialogModule.Module, frm2.DialogModule.Module, GetEditors);

                                    if (frm.Dialog.CheckSumm != frm2.Dialog.CheckSumm)
                                        report += string.Format("ЖурналРасчетов.{0}.ФормаСписка.{1}.Диалог\n", frm.Identity, item.Identity);
                                }

                                if (frm.Moxels != null)
                                    foreach (MetaItem Moxel in frm.Moxels)
                                    {
                                        MetaItem Moxel2 = frm2.Moxels.Find(x => x.Name == Moxel.Name);

                                        if (Moxel2 != null)
                                        {
                                            if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                                report += string.Format("Таблица ЖурналРасчетов.{1}.ФормаСписка.{0}.{2}\n", frm.Identity, item.Identity, Moxel.Description);
                                        }
                                        else
                                        {
                                            report += string.Format("Добавлена таблица ЖурналРасчетов.{1}.ФормаСписка.{0}.{2}\n", frm.Identity, item.Identity, Moxel.Description);
                                        }

                                    }

                            }
                            else
                                report += string.Format("Добавлена форма: ЖурналРасчетов.{0}.ФормаСписка.{1}\n", frm.Identity, item.Identity);
                        }
                    }
                    else
                        report += string.Format("Добавлен: ЖурналРасчетов.{0}\n", item.Identity);
                }
                #endregion

                //Сравним обработки
                #region Обработки
                foreach (var item in CalcVars)
                {
                    var item2 = Second.CalcVars.Find(x => x.Identity == item.Identity);

                    if (item2 != null)
                    {
                        if (item2.Form.DialogModule.CheckSumm != item.Form.DialogModule.CheckSumm)
                        {
                            if (item.Identity != "DefCls")
                            {
                                report += string.Format("Обработка.{0}.Форма.Модуль\n", item.Identity);
                                if (comparemodeules)
                                    report += CompareModules(item.Form.DialogModule.Module, item2.Form.DialogModule.Module, GetEditors);

                                if (item2.Form.Dialog.CheckSumm != item.Form.Dialog.CheckSumm)
                                    report += string.Format("Обработка.{0}.Форма.Диалог\n", item.Identity);

                                if (item.Form.Moxels != null)
                                    foreach (MetaItem Moxel in item.Form.Moxels)
                                    {
                                        MetaItem Moxel2 = item2.Form.Moxels.Find(x => x.Name == Moxel.Name);
                                        if (Moxel2 != null)
                                        {
                                            if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                                report += string.Format("Таблица Обработка.{0}.Форма.{1}\n", item.Identity, Moxel.Description);
                                        }
                                        else
                                        {
                                            report += string.Format("Добавлена таблица Обработка.{0}.Форма{1}\n", item.Identity, Moxel.Description);
                                        }

                                    }
                            }
                            else
                            {
                                string classesChange = string.Empty;

                                string[] splitter = { "\r\n" };
                                string[] classes1 = item.Form.DialogModule.Moduletext.Split(splitter, StringSplitOptions.RemoveEmptyEntries);
                                string[] classes2 = item2.Form.DialogModule.Moduletext.Split(splitter, StringSplitOptions.RemoveEmptyEntries);

                                foreach (string classesitem in classes1)
                                {
                                    if (classesitem.Substring(0, 5) == "//# {")
                                        continue;
                                    string classname = classesitem.Substring(classesitem.LastIndexOf(' ', classesitem.IndexOf('=') - 2), classesitem.IndexOf("=") - classesitem.LastIndexOf(' ', classesitem.IndexOf('=') - 2)).TrimEnd(' ').TrimStart(' ');
                                    if (!item2.Form.DialogModule.Moduletext.Contains(classname))
                                    {
                                        string calcvarname = classesitem.Substring(classesitem.IndexOf("=") + 1, classesitem.IndexOf("@") - classesitem.IndexOf("=") - 1).TrimEnd(' ').TrimStart(' ');
                                        classesChange += "    Добавлен класс " + classname + " (обработка \"" + calcvarname + "\")\n";
                                    }
                                }


                                foreach (string classesitem in classes2)
                                {
                                    if (classesitem.Substring(0, 5) == "//# {")
                                        continue;
                                    string classname = classesitem.Substring(classesitem.LastIndexOf(' ', classesitem.IndexOf('=') - 2), classesitem.IndexOf("=") - classesitem.LastIndexOf(' ', classesitem.IndexOf('=') - 2)).TrimEnd(' ').TrimStart(' ');
                                    if (!item.Form.DialogModule.Moduletext.Contains(classname))
                                        classesChange += "    Удален класс " + classname + "\n";
                                }

                                if (classesChange != string.Empty)
                                {
                                    report += item.Identity + ":\n" + classesChange;
                                }
                            }
                        }
                    }
                    else
                        report += string.Format("Добавлен: Обработка.{0}\n", item.Identity);
                }
                #endregion

                //Сравним отчеты
                #region Отчеты
                foreach (var item in Reports)
                {
                    var item2 = Second.Reports.Find(x => x.Identity == item.Identity);

                    if (item2 != null)
                    {
                        if (item2.Form.DialogModule.CheckSumm != item.Form.DialogModule.CheckSumm)
                        {
                            report += string.Format("Отчет.{0}.Форма.Модуль\n", item.Identity);
                            if (comparemodeules)
                                report += CompareModules(item.Form.DialogModule.Module, item2.Form.DialogModule.Module, GetEditors);
                            if (item2.Form.Dialog.CheckSumm != item.Form.Dialog.CheckSumm)
                                report += string.Format("Отчет.{0}.Форма.Диалог\n", item.Identity);
                        }

                        if (item.Form.Moxels != null)
                            foreach (MetaItem Moxel in item.Form.Moxels)
                            {
                                MetaItem Moxel2 = item2.Form.Moxels.Find(x => x.Name == Moxel.Name);
                                if (Moxel2 != null)
                                {
                                    if (Moxel2.CheckSumm != Moxel.CheckSumm)
                                        report += string.Format("Таблица Отчет.{0}.Форма.{1}\n", item.Identity, Moxel.Description);
                                }
                                else
                                {
                                    report += string.Format("Добавлена таблица Отчет.{0}.Форма{1}\n", item.Identity, Moxel.Description);
                                }

                            }
                    }
                    else
                        report += string.Format("Добавлен: Отчет.{0}\n", item.Identity);
                }
                #endregion
                
                return report;
            }

            private void LoadFile(string FileName)
            {
                MDFileName = FileName;
                MetadataFileTree = new MetaContainer(FileName);

                if (MetadataFileTree.Items == null)
                    return;

                MMS = MetadataFileTree.GetSubItem("Root\\Metadata\\Main MetaData Stream").Moduletext;
                GlobalModule = MetadataFileTree.GetSubItem("Root\\TypedText\\ModuleText_Number1\\MD Programm text");
                TagStream = MetadataFileTree.GetSubItem("Root\\Metadata\\TagStream");
                byte[] GUIDData_raw = MetadataFileTree.GetSubItem("Root\\Metadata\\GUIDData").data;
                byte[] buffer = new byte[16];

                while (GUIDData_raw.Length - (20 + 16 * GUIDData.Count) >= 16)
                {
                    Array.ConstrainedCopy(GUIDData_raw, 20 + 16 * GUIDData.Count, buffer, 0, 16);
                    GUIDData.Add(new Guid(buffer));
                }

                ParseMMS(MMS);
            }

            public TaskItem(string FileName)
            {
                LoadFile(FileName);
            }

            private void ParseMMS(string MMS)
            {
                int First = MMS.IndexOf("{\r\n");
                MMS = MMS.Substring(First);
                string errorlog = string.Empty;
                string[] elements = MMS.Replace("\r\n", "").Split('{');
                string[][][] objects = new string[elements.Length][][];
                int level = 0;
                int index = -1;
                int BuhFormsCount = 0;
                int i = 0;
                DescriptorType Type = DescriptorType.Null;
                MetaObject Current = null;
                string Current5 = null;
                foreach (string sub in elements)
                {
                    level++;
                    string[] ss = sub.Split("}".ToCharArray());
                    foreach (string subss in ss)
                    {
                        string[] properties = subss.Split(',');
                        string trytype = properties[0].Replace("\"", "").Replace(" ", "");

                        bool toplevel = level == 3;

                        if (toplevel)
                        {
                            if (trytype != "")
                            {
                                if (Enum.TryParse<DescriptorType>(trytype, out Type))
                                {
                                    index++;
                                    objects[index] = new string[3][];
                                    objects[index][0] = new string[1];
                                    objects[index][0][0] = Type.GetDescription();
                                    objects[index][1] = properties;
                                    objects[index][2] = new string[1];
                                    i = 0;
                                }
                                else
                                {
                                    errorlog += "DescriptorType." + trytype + ",\r\n";
                                }

                            }
                        }

                        string[] ObjParams = subss.Replace("\",\"", "|").Replace("\"", "").Split('|');

                        if (level == 4)
                        {
                            if (objects[index][2].Length <= i + 1)
                                Array.Resize<string>(ref objects[index][2], i + 1);
                            objects[index][2][i] = subss;
                            i++;
                            if (ObjParams.Length >= 4)
                            {
                                if (Type == DescriptorType.SbCnts)
                                {
                                    Subconto SB = new Subconto(ObjParams);
                                    MetaContainer FormContainer;

                                    if (SB.EditMode != 0) //Способ редактирования не "В списке"
                                    {
                                        FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\Subconto\\Subconto_Number{0}\\WorkBook", SB.ID));
                                        if (FormContainer != null)
                                            SB.Form = new Form(FormContainer);
                                    }
                                    if (SB.OneForm != 1)
                                        if (SB.LevelCount > 1)
                                        {
                                            FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\SubFolder\\SubFolder_Number{0}\\WorkBook", SB.ID));
                                            if (FormContainer != null)
                                                SB.FolderForm = new Form(FormContainer);
                                        }

                                    Subcontos.Add(SB);
                                    Current = SB;
                                }

                                if (Type == DescriptorType.Documents)
                                {
                                    Document Doc = new Document();
                                    Doc.ID = int.Parse(ObjParams[0]);
                                    Doc.Identity = ObjParams[1];
                                    Doc.Alias = ObjParams[2];
                                    MetaContainer FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\Document\\Document_Number{0}\\WorkBook", Doc.ID));
                                    if (FormContainer != null)
                                        Doc.Form = new Form(FormContainer);
                                    Doc.TransactionModule = MetadataFileTree.GetSubItem(String.Format("Root\\TypedText\\Transact_Number{0}\\MD Programm text", Doc.ID));
                                    Documents.Add(Doc);
                                    Current = Doc;
                                }

                                if (Type == DescriptorType.Buh)
                                {
                                    Buh = new BuhParameters();
                                    Buh.ID = int.Parse(ObjParams[0]);
                                    Buh.Identity = ObjParams[1];
                                    Buh.Alias = ObjParams[2];

                                    Document Doc = Documents.Find(x => x.Identity == "Операция");
                                    MetaContainer FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\Operation\\Operation_Number{0}\\WorkBook", ObjParams[0]));
                                    if (FormContainer != null)
                                        Doc.Form = new Form(FormContainer);

                                    FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\AccountChart\\AccountChart_Number{0}\\WorkBook", Buh.ID));

                                    if (FormContainer != null)
                                        Buh.AccountChart.Add(new Form(FormContainer));

                                    Current = Doc;
                                }

                                if (Type == DescriptorType.Journalisters)
                                {
                                    Journal obj = new Journal();
                                    obj.ID = int.Parse(ObjParams[0]);
                                    obj.Identity = ObjParams[1];
                                    obj.Alias = ObjParams[2];
                                    Journals.Add(obj);
                                    Current = obj;
                                }

                                if (Type == DescriptorType.CJ)
                                {
                                    CalcJournal obj = new CalcJournal();
                                    obj.ID = int.Parse(ObjParams[0]);
                                    obj.Identity = ObjParams[1];
                                    obj.Alias = ObjParams[2];
                                    CJ.Add(obj);
                                    Current = obj;
                                }

                                if (Type == DescriptorType.ReportList)
                                {
                                    ReportItem obj = new ReportItem();

                                    obj.ID = int.Parse(ObjParams[0]);
                                    obj.Identity = ObjParams[1];
                                    obj.Alias = ObjParams[2];
                                    obj.Form = new Form(MetadataFileTree.GetSubCatalog(String.Format("Root\\Report\\Report_Number{0}\\WorkBook", obj.ID)));
                                    Reports.Add(obj);
                                    Current = obj;
                                }

                                if (Type == DescriptorType.CalcVars)
                                {
                                    CalcVarItem obj = new CalcVarItem();

                                    obj.ID = int.Parse(ObjParams[0]);
                                    obj.Identity = ObjParams[1];
                                    obj.Alias = ObjParams[2];
                                    obj.Form = new Form(MetadataFileTree.GetSubCatalog(String.Format("Root\\CalcVar\\CalcVar_Number{0}\\WorkBook", obj.ID)));
                                    CalcVars.Add(obj);
                                    Current = obj;
                                }

                                if (Type == DescriptorType.Algorithms)
                                {
                                    CalculationAlgorithm obj = new CalculationAlgorithm();

                                    obj.ID = int.Parse(ObjParams[0]);
                                    obj.Identity = ObjParams[1];
                                    obj.Alias = ObjParams[2];
                                    obj.CalculationModule = MetadataFileTree.GetSubItem(String.Format("Root\\TypedText\\CalcAlg_Number{0}\\MD Programm text", obj.ID));
                                    Algorithms.Add(obj);
                                    Current = obj;
                                }
                            }
                        }

                        if (level == 5)
                        {
                            Current5 = trytype;

                            if (Type == DescriptorType.Buh && Current5 == "Form")
                                BuhFormsCount++;
                        }

                        if (level == 7)
                        {
                            Current5 = trytype;
                            if (Type == DescriptorType.Buh && trytype == "Form")
                                BuhFormsCount++;
                        }

                        if ((level == 6 || level == 8) && Current5 == "Form" && ObjParams.Length == 4)
                        {
                            MetaContainer FormContainer = null;

                            if (Type == DescriptorType.SbCnts)
                                FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\SubList\\SubList_Number{0}\\WorkBook", ObjParams[0]));

                            if (Type == DescriptorType.Journalisters)
                                FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\Journal\\Journal_Number{0}\\WorkBook", ObjParams[0]));

                            if (Type == DescriptorType.CJ)
                                FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\CalcJournal\\CalcJournal_Number{0}\\WorkBook", ObjParams[0]));

                            if (Type == DescriptorType.Buh)
                            {
                                switch (BuhFormsCount)
                                {
                                    case 1:            //Форма списка плана счетов
                                        FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\AccountChartList\\AccountChartList_Number{0}\\WorkBook", ObjParams[0]));
                                        break;
                                    case 2:            //ХЗ
                                        FormContainer = null;//MetadataFileTree.GetSubCatalog(String.Format("Root\\AccountChart\\AccountChart_Number{0}\\WorkBook", Buh.ID));
                                        break;
                                    case 3:            //Форма списка проводок
                                        FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\ProvList\\ProvList_Number{0}\\WorkBook", ObjParams[0]));
                                        break;
                                    case 4:            //Форма списка операции
                                        FormContainer = MetadataFileTree.GetSubCatalog(String.Format("Root\\OperationList\\OperationList_Number{0}\\WorkBook", ObjParams[0]));
                                        break;
                                }
                            }

                            Form frm = new Form();

                            if (FormContainer != null)
                                frm = new Form(FormContainer);

                            frm.ID = int.Parse(ObjParams[0]);
                            frm.Identity = ObjParams[1];
                            frm.Alias = ObjParams[2];
                            frm.Description = ObjParams[3];

                            if (Type == DescriptorType.SbCnts)
                                ((Subconto)Current).ListForms.Add(frm);

                            if (Type == DescriptorType.Journalisters)
                                ((Journal)Current).ListForms.Add(frm);

                            if (Type == DescriptorType.CJ)
                                ((CalcJournal)Current).ListForms.Add(frm);

                            if (Type == DescriptorType.Buh)
                            {
                                switch (BuhFormsCount)
                                {
                                    case 1:            //Форма списка плана счетов
                                        Buh.AccountChartList.Add(frm);
                                        break;
                                    case 2:            //ХЗ
                                        Buh.AccountChart.Add(frm);
                                        break;
                                    case 3:            //Форма списка проводок
                                        Buh.ProvListList.Add(frm);
                                        break;
                                    case 4:            //Форма списка операции
                                        Buh.OperationList.Add(frm);
                                        break;
                                }
                            }

                        }
                    }
                    level -= ss.Length - 1;
                }

                Array.Resize<string[][]>(ref objects, index + 1);
                if (errorlog.Length > 1)
                    File.WriteAllText(MDFileName + ".errorlog", errorlog);
            }
        }

        public class MetaData
        {
            MetaObject MainDataContDef;

        }
        #endregion

        #region Интерфейсы OLE
        [ComImport]
        [Guid("0000000d-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IEnumSTATSTG
        {
            // The user needs to allocate an STATSTG array whose size is celt.
            [PreserveSig]
            uint Next(
                uint celt,
                [MarshalAs(UnmanagedType.LPArray), Out]
                    System.Runtime.InteropServices.ComTypes.STATSTG[] rgelt,
                out uint pceltFetched
            );
            void Skip(uint celt);
            void Reset();
            [return: MarshalAs(UnmanagedType.Interface)]
            IEnumSTATSTG Clone();
        }

        [ComImport]
        [Guid("0000000b-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IStorage
        {
            void CreateStream(
                /* [string][in] */ string pwcsName,
                /* [in] */ uint grfMode,
                /* [in] */ uint reserved1,
                /* [in] */ uint reserved2,
                /* [out] */ out IStream ppstm);

            void OpenStream(
                /* [string][in] */ string pwcsName,
                /* [unique][in] */ IntPtr reserved1,
                /* [in] */ uint grfMode,
                /* [in] */ uint reserved2,
                /* [out] */ out IStream ppstm);

            void CreateStorage(
                /* [string][in] */ string pwcsName,
                /* [in] */ uint grfMode,
                /* [in] */ uint reserved1,
                /* [in] */ uint reserved2,
                /* [out] */ out IStorage ppstg);

            void OpenStorage(
                /* [string][unique][in] */ string pwcsName,
                /* [unique][in] */ IStorage pstgPriority,
                /* [in] */ uint grfMode,
                /* [unique][in] */ IntPtr snbExclude,
                /* [in] */ uint reserved,
                /* [out] */ out IStorage ppstg);

            void CopyTo(
                /* [in] */ uint ciidExclude,
                /* [size_is][unique][in] */ Guid rgiidExclude, // should this be an array?
                /* [unique][in] */ IntPtr snbExclude,
                /* [unique][in] */ IStorage pstgDest);

            void MoveElementTo(
                /* [string][in] */ string pwcsName,
                /* [unique][in] */ IStorage pstgDest,
                /* [string][in] */ string pwcsNewName,
                /* [in] */ uint grfFlags);

            void Commit(
                /* [in] */ STGC grfCommitFlags);

            void Revert();

            void EnumElements(
                /* [in] */ uint reserved1,
                /* [size_is][unique][in] */ IntPtr reserved2,
                /* [in] */ uint reserved3,
                /* [out] */ out IEnumSTATSTG ppenum);

            void DestroyElement(
                /* [string][in] */ string pwcsName);

            void RenameElement(
                /* [string][in] */ string pwcsOldName,
                /* [string][in] */ string pwcsNewName);

            void SetElementTimes(
                /* [string][unique][in] */ string pwcsName,
                /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pctime,
                /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME patime,
                /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pmtime);

            void SetClass(
                /* [in] */ Guid clsid);

            void SetStateBits(
                /* [in] */ uint grfStateBits,
                /* [in] */ uint grfMask);

            void Stat(
                /* [out] */ out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg,
                /* [in] */ uint grfStatFlag);

        }

        [Flags]
        private enum STGC : int
        {
            DEFAULT = 0,
            OVERWRITE = 1,
            ONLYIFCURRENT = 2,
            DANGEROUSLYCOMMITMERELYTODISKCACHE = 4,
            CONSOLIDATE = 8
        }

        [Flags]
        private enum STGM : int
        {
            DIRECT = 0x00000000,
            TRANSACTED = 0x00010000,
            SIMPLE = 0x08000000,
            READ = 0x00000000,
            WRITE = 0x00000001,
            READWRITE = 0x00000002,
            SHARE_DENY_NONE = 0x00000040,
            SHARE_DENY_READ = 0x00000030,
            SHARE_DENY_WRITE = 0x00000020,
            SHARE_EXCLUSIVE = 0x00000010,
            PRIORITY = 0x00040000,
            DELETEONRELEASE = 0x04000000,
            NOSCRATCH = 0x00100000,
            CREATE = 0x00001000,
            CONVERT = 0x00020000,
            FAILIFTHERE = 0x00000000,
            NOSNAPSHOT = 0x00200000,
            DIRECT_SWMR = 0x00400000,
        }

        [Flags]
        private enum STATFLAG : uint
        {
            STATFLAG_DEFAULT = 0,
            STATFLAG_NONAME = 1,
            STATFLAG_NOOPEN = 2
        }

        [Flags]
        private enum STGTY : int
        {
            STGTY_STORAGE = 1,
            STGTY_STREAM = 2,
            STGTY_LOCKBYTES = 3,
            STGTY_PROPERTY = 4
        }

        //Читает IStream в массив байт
        private static byte[] ReadIStream(IStream pIStream)
        {
            System.Runtime.InteropServices.ComTypes.STATSTG StreamInfo;
            pIStream.Stat(out StreamInfo, 0);
            byte[] data = new byte[StreamInfo.cbSize];
            pIStream.Read(data, (int)StreamInfo.cbSize, IntPtr.Zero);
            return data;
        }


        [ComVisible(false)]
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000A-0000-0000-C000-000000000046")]
        public interface ILockBytes
        {
            void ReadAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbRead);
            void WriteAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbWritten);
            void Flush();
            void SetSize(long cb);
            void LockRegion(long libOffset, long cb, int dwLockType);
            void UnlockRegion(long libOffset, long cb, int dwLockType);
            void Stat(out System.Runtime.InteropServices.STATSTG pstatstg, int grfStatFlag);

        }

        class NativeMethods
        {
            [DllImport("ole32.dll")]
            public static extern int StgIsStorageFile(
                [MarshalAs(UnmanagedType.LPWStr)] string pwcsName);

            [DllImport("ole32.dll")]
            public static extern int StgOpenStorage(
                [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
                IStorage pstgPriority,
                STGM grfMode,
                IntPtr snbExclude,
                uint reserved,
                out IStorage ppstgOpen);

            [DllImport("ole32.dll")]
            public static extern int StgCreateDocfile(
                [MarshalAs(UnmanagedType.LPWStr)]string pwcsName,
                STGM grfMode,
                uint reserved,
                out IStorage ppstgOpen);

            [DllImport("ole32.dll")]
            static extern int StgOpenStorageOnILockBytes(ILockBytes plkbyt,
               IStorage pStgPriority, uint grfMode, IntPtr snbEnclude, uint reserved,
               out IStorage ppstgOpen);

            [DllImport("ole32.dll")]
            public extern static int CreateILockBytesOnHGlobal(IntPtr hGlobal, [MarshalAs(UnmanagedType.Bool)] bool fDeleteOnRelease, out ILockBytes ppLkbyt);
        }

        #endregion
    }
}
