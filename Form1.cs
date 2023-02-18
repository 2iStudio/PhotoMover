using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Dialogs;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using System.Windows.Forms;

namespace PhotoMover
{

    public partial class Form1 : Form
    {
        //[System.Runtime.InteropServices.DllImport("kernel32.dll")]
        //private static extern bool AllocConsole();

        private bool _mouseDownFlg = false;
        private PointF _oldPoint;
        private System.Drawing.Drawing2D.Matrix _matAffine;
        private Graphics _gPic;
        private Bitmap _image;

        private string _saveFolderPass;
        private string _saveFilePass;

        List<ShowImage> _images = new List<ShowImage>();
        List<TPSet> _TPSets = new List<TPSet>();
        int nowLoading = 0;

        public int _fitMode = 0;

        int _CopyOption1 = 0;
        int _CopyOption2 = 0;
        int _CopyOption3 = 0;

        bool copying = false;
        bool isAutoSave = false;

        string defaultTempCSV;

        const string integrityString = "o54WVn14ITxKMcS2rYsd";

        /// <summary>
        /// csvファイル:1行目に読み込めるかチェックする文字列,作業フォルダの絶対パス、最後に見ていたファイルの絶対パス,tpFolderのパス*10
        /// 以降0番目から,絶対ファイルパス,tpBool * 10, isCheckedを1行ずつ(ファイル名は絶対ファイルパスから取得できるので省略)
        /// 
        /// AffinImage class はHowToUse.pdfに記載のサイトより
        /// </summary>

        public Form1()
        {
            InitializeComponent();


            CreateGraphics(pictureBox1, ref _gPic);
            pictureBox1.AllowDrop = true;
            _matAffine = new System.Drawing.Drawing2D.Matrix();
            this.pictureBox1.MouseWheel += new System.Windows.Forms.MouseEventHandler(this.pictureBox1_MouseWheel);

            //AllocConsole();

            this.KeyPreview = true;
            MakeTPSet();

            CopyProgressBar.Visible = false;

            //コピー設定
            _CopyOption1 = Properties.Settings.Default.option1;
            _CopyOption2 = Properties.Settings.Default.option2;
            _CopyOption3 = Properties.Settings.Default.option3;

            CopyOptionComboBox1.SelectedIndex = _CopyOption1;
            CopyOptionComboBox2.SelectedIndex = _CopyOption2;
            CopyOptionComboBox3.SelectedIndex = _CopyOption3;

            //画像ズーム設定
            _fitMode = Properties.Settings.Default.zoomOption;
            PicDisplayComboBox.SelectedIndex = _fitMode;

            //復元設定
            RestoreCheckBox.Checked = Properties.Settings.Default.loadSettings;

            //オートセーブ設定
            isAutoSave = Properties.Settings.Default.autoSave;
            AutoSaveCheckBox.Checked = isAutoSave;

            

            _saveFolderPass = System.IO.Path.GetDirectoryName(Application.ExecutablePath);

            defaultTempCSV = System.IO.Path.Combine(_saveFolderPass, "temp.csv");

            _saveFolderPass = System.IO.Path.Combine(_saveFolderPass, "SaveData");


            if (!File.Exists(_saveFolderPass)) Directory.CreateDirectory(_saveFolderPass);

            if (Properties.Settings.Default.loadSettings)
            {
                if(File.Exists(Properties.Settings.Default.workFile))
                {
                    try
                    {
                        LoadSaveFolder(Properties.Settings.Default.workFile);
                    }
                    catch
                    {

                    }
                }
            }
        }

        //---------------データ保存-----------------


        void LoadSaveFolder(string path)
        {
            using(StreamReader _workfileRead = new StreamReader(@path))
            {

                //1行目:以前の作業ファイルのパスと見ていたファイルを拾う
                string line = _workfileRead.ReadLine();
                string[] values = line.Split('|');

                if (values[0] == integrityString)
                {
                    workFolderText.Text = values[1];
                    string preLoaded = values[2];

                    for (int i = 3; i < 13; i++)
                    {
                        /*
                        if (File.Exists(values[i]))
                        {
                            Console.WriteLine(values[i] + "exists");

                            Console.WriteLine(_TPSets[i - 2]._textBox.Text + "has load");
                        }
                        else
                        {
                            Console.WriteLine(values[i] + "notexists");
                            _TPSets[i - 2]._textBox.Text = "def";
                        }*/

                        _TPSets[i - 3]._textBox.Text = values[i];
                    }


                    int preLoadedIndex = 0;

                    _images.Clear();
                    ShowImage tempShowImage;

                    while (!_workfileRead.EndOfStream)
                    {
                        line = _workfileRead.ReadLine();
                        values = line.Split('|');


                        tempShowImage = new ShowImage(values[0].ToString());

                        for (int i = 1; i < 11; i++)
                        {
                            if (int.Parse(values[i].ToString()) == 1)
                            {
                                tempShowImage.tpBool[i - 1] = true;
                            }
                        }

                        if (int.Parse(values[11].ToString()) == 1)
                        {
                            tempShowImage.isChecked = true;
                        }

                        _images.Add(tempShowImage);
                    }

                    List<ShowImage> deleteImage = new List<ShowImage>();

                    for (int i = 0; i < _images.Count; i++)
                    {
                        if (!File.Exists(_images[i]._filePath))
                        {
                            deleteImage.Add(_images[i]);
                        }
                    }

                    foreach (ShowImage si in deleteImage)
                    {
                        _images.Remove(si);
                    }

                    foreach (ShowImage si in _images)
                    {
                        if (si._filePath == preLoaded) break;
                        else preLoadedIndex++;
                    }
                    if (preLoadedIndex == _images.Count)
                    {
                        preLoadedIndex = 0;
                    }

                    LoadImage(preLoadedIndex);
                }

                ActiveControl = null;

                imageNumBox.Text = (nowLoading + 1).ToString();
                imageNumBoxConst.Text = "/" + _images.Count.ToString();

                _saveFilePass = path;
                Properties.Settings.Default.workFile = _saveFilePass;
                Properties.Settings.Default.Save();

                CSVtextbox.Text = System.IO.Path.GetFileName(path);

                pictureBox1.Focus();
            }
        }

        void SaveWorks()
        {
            DateTime start = DateTime.Now;
            if (_saveFilePass == null | _images.Count == 0) return;

            using(StreamWriter _workfileWrite = new StreamWriter(@_saveFilePass, false))
            {
                try
                {
                    _workfileWrite.Write(integrityString);
                    _workfileWrite.Write("|");

                    _workfileWrite.Write(workFolderText.Text);
                    _workfileWrite.Write("|");
                    _workfileWrite.Write(_images[nowLoading]._filePath);

                    for(int i = 0; i < 10; i++)
                    {
                        _workfileWrite.Write("|");
                        _workfileWrite.Write(_TPSets[i]._textBox.Text);
                    }

                    _workfileWrite.Write("\r\n");


                    foreach (var si in _images)
                    {
                        string[] line = new string[12];

                        line[0] = si._filePath;

                        for (int i = 0; i < 10; i++)
                        {
                            if (si.tpBool[i]) line[i + 1] = "|1";
                            else line[i + 1] = "|0";
                        }

                        if (si.isChecked) line[11] = "|1";
                        else line[11] = "|0";

                        foreach (string text in line)
                        {
                            _workfileWrite.Write(text);
                        }


                        _workfileWrite.Write("\r\n");
                    }

                    Properties.Settings.Default.workFile = _saveFilePass;
                    Properties.Settings.Default.Save();

                    CSVtextbox.Text = System.IO.Path.GetFileName(_saveFilePass);
                }
                catch (Exception e)
                {
                }
            }
            
        }

        //-------------コンポーネント周り---------------

        //--------------キー設定-------------------

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (copying) return;

            if (e.KeyCode == Keys.Left | e.KeyCode == Keys.A)
            {
                ChangeImage(false);
            }
            else if (e.KeyCode == Keys.Right | e.KeyCode == Keys.D)
            {
                ChangeImage(true);
            }
            else if(e.KeyCode == Keys.NumPad1 | e.KeyCode == Keys.D1)
            {
                CheckBoxControl(0);
            }
            else if (e.KeyCode == Keys.NumPad2 | e.KeyCode == Keys.D2)
            {
                CheckBoxControl(1);
            }
            else if (e.KeyCode == Keys.NumPad3 | e.KeyCode == Keys.D3)
            {
                CheckBoxControl(2);
            }
            else if (e.KeyCode == Keys.NumPad4 | e.KeyCode == Keys.D4)
            {
                CheckBoxControl(3);
            }
            else if (e.KeyCode == Keys.NumPad5 | e.KeyCode == Keys.D5)
            {
                CheckBoxControl(4);
            }
            else if (e.KeyCode == Keys.NumPad6 | e.KeyCode == Keys.D6)
            {
                CheckBoxControl(5);
            }
            else if (e.KeyCode == Keys.NumPad7 | e.KeyCode == Keys.D7)
            {
                CheckBoxControl(6);
            }
            else if (e.KeyCode == Keys.NumPad8 | e.KeyCode == Keys.D8)
            {
                CheckBoxControl(7);
            }
            else if (e.KeyCode == Keys.NumPad9 | e.KeyCode == Keys.D9)
            {
                CheckBoxControl(8);
            }
            else if (e.KeyCode == Keys.NumPad0 | e.KeyCode == Keys.D0)
            {
                CheckBoxControl(9);
            }
            else if(e.KeyCode == Keys.S && e.Control == true)
            {
                SaveButtonFunction();
            }
        }

        private void CheckBoxControl(int i)
        {
            if (copying) return;
            if (_images.Count == 0) return;
            _images[nowLoading].tpBool[i] = !_images[nowLoading].tpBool[i];
            _TPSets[i]._checkBox.Checked = _images[nowLoading].tpBool[i];

            if(_TPSets[i]._checkBox.Checked)
            {
                _TPSets[i]._textBox.BackColor = SystemColors.GradientActiveCaption;
            }
            else
            {
                _TPSets[i]._textBox.BackColor = Color.White;
            }

            if (isAutoSave) SaveWorks();
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            if (pictureBox1.Image == null) return;

            CreateGraphics(pictureBox1, ref _gPic);
            pictureBox1.AllowDrop = true;
            _matAffine = new System.Drawing.Drawing2D.Matrix();

            

            if(copying)
            {
                _image = Properties.Resources.copying;

                _matAffine.ZoomFit(pictureBox1, _image, _fitMode);

                // 画像の描画
                RedrawImage();

                loadingImageNameLabel.Text = "copying";

                pictureBox1.Focus();
            }
            else
            {
                LoadImage(nowLoading);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_images.Count != 0)
            {
                if (System.Globalization.CultureInfo.CurrentCulture.Name == "ja-JP")
                {
                    //メッセージボックスを表示する
                    DialogResult result = MessageBox.Show("現在の作業を保存してから終了しますか？",
                        "閉じる",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Exclamation,
                        MessageBoxDefaultButton.Button1);

                    //何が選択されたか調べる
                    if (result == DialogResult.Yes)
                    {
                        SaveButtonFunction();
                    }
                    else if (result == DialogResult.No)
                    {
                        //「いいえ」が選択された時
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        //「キャンセル」が選択された時
                        e.Cancel = true;
                    }
                }
                else
                {
                    //メッセージボックスを表示する
                    DialogResult result = MessageBox.Show("Save your current work before exiting?",
                        "Close",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Exclamation,
                        MessageBoxDefaultButton.Button1);

                    //何が選択されたか調べる
                    if (result == DialogResult.Yes)
                    {
                        SaveButtonFunction();
                    }
                    else if (result == DialogResult.No)
                    {
                        //「いいえ」が選択された時
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        //「キャンセル」が選択された時
                        e.Cancel = true;
                    }
                }
            }
        }

        void ChangeImage(bool next)
        {
            if (copying) return;
            if (next)
            {
                if (_images.Count == 0)
                {
                    return;
                }
                else if (nowLoading == _images.Count-1)
                {
                    nowLoading = 0;
                    LoadImage(nowLoading);
                }
                else
                {
                    nowLoading++;
                    LoadImage(nowLoading);
                }
            }
            else
            {
                if (_images.Count == 0)
                {
                    return;
                }
                else if (nowLoading == 0)
                {
                    nowLoading = _images.Count - 1;
                    LoadImage(nowLoading);
                }
                else
                {
                    nowLoading--;
                    LoadImage(nowLoading);
                }
            }

            if (isAutoSave) SaveWorks();

            pictureBox1.Focus();
        }

        private void PicDisplayComboBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            _fitMode = PicDisplayComboBox.SelectedIndex;

            Properties.Settings.Default.zoomOption = PicDisplayComboBox.SelectedIndex;
            Properties.Settings.Default.Save();

            LoadImage(nowLoading);
        }


        private void CopyOptionComboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (copying) return;
            _CopyOption1 = CopyOptionComboBox1.SelectedIndex;
            Properties.Settings.Default.option1 = _CopyOption1;
            Properties.Settings.Default.Save();
        }


        private void CopyOptionComboBox2_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (copying) return;
            _CopyOption2 = CopyOptionComboBox2.SelectedIndex;
            Properties.Settings.Default.option2 = _CopyOption2;
            Properties.Settings.Default.Save();
        }


        private void CopyOptionComboBox3_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (copying) return;
            _CopyOption3 = CopyOptionComboBox3.SelectedIndex;
            Properties.Settings.Default.option3 = _CopyOption3;
            Properties.Settings.Default.Save();
        }


        private void SaveButton_Click(object sender, EventArgs e)
        {
            SaveButtonFunction();
        }

        void SaveButtonFunction()
        {
            if (_images.Count == 0) return;

            if (_saveFilePass == null | !File.Exists(_saveFilePass) | _saveFilePass == defaultTempCSV)
            {
                //SaveFileDialogクラスのインスタンスを作成
                SaveFileDialog sfd = new SaveFileDialog();
                //はじめのファイル名を指定する
                sfd.FileName = "work.csv";

                if (_saveFilePass != null)
                {
                    try
                    {
                        sfd.FileName = System.IO.Path.GetFileName(_saveFilePass);
                    }
                    catch
                    {
                        sfd.FileName = "work.csv";
                    }
                }
                else if (workFolderText.Text != null)
                {
                    try
                    {
                        sfd.FileName = System.IO.Path.GetFileName(workFolderText.Text) + ".csv";
                    }
                    catch
                    {
                        sfd.FileName = "work.csv";
                    }
                }

                //デフォルトディレクトリ設定
                sfd.InitialDirectory = @_saveFolderPass;
                //[ファイルの種類]に表示される選択肢を指定する
                sfd.Filter = "csvファイル(*.csv)|*.csv";
                sfd.Title = "名前を付けて保存";

                if (System.Globalization.CultureInfo.CurrentCulture.Name != "ja-JP")
                {
                    sfd.Title = "Save as";
                }

                //ダイアログを表示する
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    _saveFilePass = sfd.FileName;
                    SaveWorks();
                }
            }
            else
            {
                SaveWorks();
            }
        }

        private void SaveAsButton_Click(object sender, EventArgs e)
        {
            if (_images.Count == 0) return;

            //SaveFileDialogクラスのインスタンスを作成
            SaveFileDialog sfd = new SaveFileDialog();
            //はじめのファイル名を指定する
            sfd.FileName = "work.csv";

            if(_saveFilePass !=null)
            {
                try
                {
                    sfd.FileName = System.IO.Path.GetFileName(_saveFilePass);
                }
                catch
                {
                    sfd.FileName = "work.csv";
                }
            }
            else if(workFolderText.Text != null)
            {
                try
                {
                    sfd.FileName = System.IO.Path.GetFileName(workFolderText.Text) + ".csv";
                }
                catch
                {
                    sfd.FileName = "work.csv";
                }
            }

            //デフォルトディレクトリ設定
            sfd.InitialDirectory = @_saveFolderPass;
            //[ファイルの種類]に表示される選択肢を指定する
            sfd.Filter = "csvファイル(*.csv)|*.csv";
            sfd.Title = "名前を付けて保存";

            if (System.Globalization.CultureInfo.CurrentCulture.Name != "ja-JP")
            {
                sfd.Title = "Save as";
            }

            //ダイアログを表示する
            if (sfd.ShowDialog() == DialogResult.OK)
            {

                _saveFilePass = sfd.FileName;
                SaveWorks();
            }
        }

        private void LoadCSVButton_Click(object sender, EventArgs e)
        {
            //SaveFileDialogクラスのインスタンスを作成
            OpenFileDialog ofDialog = new OpenFileDialog();
            //はじめのファイル名を指定する
            ofDialog.FileName = "";

            //デフォルトディレクトリ設定
            ofDialog.InitialDirectory = @_saveFolderPass;
            //[ファイルの種類]に表示される選択肢を指定する
            ofDialog.Filter = "csvファイル(*.csv)|*.csv";
            ofDialog.Title = "読み込むファイルを指定してください";

            if(System.Globalization.CultureInfo.CurrentCulture.Name != "ja-JP")
            {
                ofDialog.Title = "Select csv to load";
            }

            //ダイアログを表示する
            if (ofDialog.ShowDialog() == DialogResult.OK)
            {
                _saveFilePass = ofDialog.FileName;
                LoadSaveFolder(ofDialog.FileName);
            }
        }


        private void RestoreCheckBox_Click(object sender, EventArgs e)
        {
            
            Properties.Settings.Default.loadSettings = RestoreCheckBox.Checked;

            if(RestoreCheckBox.Checked)
            {
                if (!File.Exists(defaultTempCSV))
                {
                    if (_saveFilePass == null)
                    {
                        _saveFilePass = defaultTempCSV;
                        SaveWorks();
                    }
                }
            }

            Properties.Settings.Default.Save();
        }

        private void AutoSaveCheckBox_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.autoSave = AutoSaveCheckBox.Checked;
            isAutoSave = AutoSaveCheckBox.Checked;
            Properties.Settings.Default.Save();

        }

        private void AutoSaveCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (_images.Count == 0) return;

            if (!File.Exists(defaultTempCSV))
            {
                if(_saveFilePass == null)
                {
                    _saveFilePass = defaultTempCSV;
                    SaveWorks();
                }
            }
        }

        private void CSVtextbox_DragDrop(object sender, DragEventArgs e)
        {
            if (copying) return;

            string[] sFileName = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            if (sFileName.Length <= 0)
            {
                return;
            }

            try
            {
                LoadSaveFolder(sFileName[0]);
            }
            catch
            {
            }
        }

        private void CSVtextbox_DragEnter(object sender, DragEventArgs e)
        {
            if (copying) return;

            string[] sFileName = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            if (System.IO.Path.GetExtension(sFileName[0]) != ".csv") return;

            e.Effect = DragDropEffects.All;
        }

        private void tpFolderCheckBox_Click(object sender, EventArgs e)
        {
            if (copying) return;
            int i = 0;

            CheckBox checkBox = sender as CheckBox;

            foreach(TPSet tP in _TPSets)
            {
                if (tP._checkBox.Name == checkBox.Name) break;
                else i++;
            }

            if (i >= _TPSets.Count) return;

            CheckBoxControl(i);
        }

        private void tpFolderDeleteButton_Click(object sender, EventArgs e)
        {
            if (copying) return;
            int i = 0;

            Button button = sender as Button;

            foreach (TPSet tP in _TPSets)
            {
                if (tP._deleteButton.Name == button.Name) break;
                else i++;
            }

            if (i >= _TPSets.Count) return;

            _TPSets[i]._textBox.Text = "";

            if(isAutoSave) SaveWorks();
        }

        private void SelectTPFolder(object sender, EventArgs e)
        {
            if (copying) return;

            Button button = sender as Button;

            int i = 0;

            foreach (TPSet tP in _TPSets)
            {
                if (button.Name == tP._button.Name) break;
                else i++;
            }

            

            if(System.Globalization.CultureInfo.CurrentCulture.Name == "ja-JP")
            {
                using (var file_dlg = new CommonOpenFileDialog()
                {

                    Title = "転送先フォルダを選択してください"
    ,
                    InitialDirectory = @"C:\"
    ,
                    IsFolderPicker = true,
                })
                {
                    if (file_dlg.ShowDialog() != CommonFileDialogResult.Ok)
                    {
                        return;
                    }
                    _TPSets[i]._textBox.Text = file_dlg.FileName;
                }
            }
            else
            {
                using (var file_dlg = new CommonOpenFileDialog()
                {

                    Title = "Select the destination folder"
    ,
                    InitialDirectory = @"C:\"
    ,
                    IsFolderPicker = true,
                })
                {
                    if (file_dlg.ShowDialog() != CommonFileDialogResult.Ok)
                    {
                        return;
                    }
                    _TPSets[i]._textBox.Text = file_dlg.FileName;
                }
            }
        }
        private void workFolderButton_Click(object sender, EventArgs e)
        {
            if (copying) return;

            
            if(System.Globalization.CultureInfo.CurrentCulture.Name == "ja-JP")
            {
                using (var file_dlg = new CommonOpenFileDialog()
                {
                    Title = "転送先フォルダを選択してください"
    ,
                    InitialDirectory = @"C:\"
    ,
                    IsFolderPicker = true,
                })
                {
                    if (file_dlg.ShowDialog() != CommonFileDialogResult.Ok)
                    {
                        return;
                    }
                    workFolderText.Text = file_dlg.FileName;
                }
            }
            else
            {
                using (var file_dlg = new CommonOpenFileDialog()
                {
                    Title = "Select the destination folder"
    ,
                    InitialDirectory = @"C:\"
    ,
                    IsFolderPicker = true,
                })
                {
                    if (file_dlg.ShowDialog() != CommonFileDialogResult.Ok)
                    {
                        return;
                    }
                    workFolderText.Text = file_dlg.FileName;
                }
            }

            GetImages();

            ActiveControl = null;

            imageNumBox.Text = (nowLoading + 1).ToString();
            imageNumBoxConst.Text = "/" + _images.Count.ToString();

            pictureBox1.Focus();
        }


        private void Text_DragEnter(object sender, DragEventArgs e)
        {
            if (copying) return;

            string[] sFileName = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            FileAttributes attr1 = File.GetAttributes(sFileName[0]);

            if ((attr1 & FileAttributes.Directory) != FileAttributes.Directory)
            {
                return;
            }

            e.Effect = DragDropEffects.All;
        }

        private void Text_Drop(object sender, DragEventArgs e)
        {
            if (copying) return;

            //ドロップされたファイルの一覧を取得
            string[] sFileName = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            if (sFileName.Length <= 0)
            {
                return;
            }

            // ドロップ先がTextBoxであるかチェック
            TextBox TargetTextBox = sender as TextBox;

            if (TargetTextBox == null)
            {
                // TextBox以外のためイベントを何もせずイベントを抜ける。
                return;
            }

            // 現状のTextBox内のデータを削除
            TargetTextBox.Text = "";

            // TextBoxドラックされた文字列を設定
            TargetTextBox.Text = sFileName[0]; // 配列の先頭文字列を設定

            ActiveControl = null;

            imageNumBox.Text = (nowLoading + 1).ToString();
            imageNumBoxConst.Text = "/" + _images.Count.ToString();

            pictureBox1.Focus();
        }

        private void workFolderText_DragEnter(object sender, DragEventArgs e)
        {
            if (copying) return;

            e.Effect = DragDropEffects.All;
        }

        private void workFolderText_Drop(object sender, DragEventArgs e)
        {
            if (copying) return;

            //ドロップされたファイルの一覧を取得
            string[] sFileName = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            if (sFileName.Length <= 0)
            {
                return;
            }

            // ドロップ先がTextBoxであるかチェック
            TextBox TargetTextBox = sender as TextBox;

            if (TargetTextBox == null)
            {
                // TextBox以外のためイベントを何もせずイベントを抜ける。
                return;
            }

            // 現状のTextBox内のデータを削除
            TargetTextBox.Text = "";

            // TextBoxドラックされた文字列を設定
            TargetTextBox.Text = sFileName[0]; // 配列の先頭文字列を設定

            FileAttributes attr1 = File.GetAttributes(sFileName[0]);

            int i = 0;

            workFolderText.Text = "";


            if ((attr1 & FileAttributes.Directory) != FileAttributes.Directory)
            {
                workFolderText.Text = System.IO.Path.GetDirectoryName(sFileName[0]);
                GetImages();

                foreach (ShowImage si in _images)
                {
                    if (si._filePath == sFileName[0]) break;
                    else i++;
                }
            }
            else
            {
                workFolderText.Text = sFileName[0];
                GetImages();
            }

            LoadImage(i);

            ActiveControl = null;

            imageNumBox.Text = (nowLoading + 1).ToString();
            imageNumBoxConst.Text = "/" + _images.Count.ToString();

            pictureBox1.Focus();
        }

        void GetImages()
        {
            if (copying) return;

            if (workFolderText.Text == null) return;

            _saveFilePass = defaultTempCSV;
            

            _images.Clear();

            IEnumerable<string> files = Directory.EnumerateFiles(@workFolderText.Text, "*", SearchOption.AllDirectories);

            foreach (string f in files)
            {
                _images.Add(new ShowImage(f));
            }
        }

        void MakeTPSet()
        {
            _TPSets.Clear();
            _TPSets.Add(new TPSet(tpFolderCheckBox1, tpFolderButton1, tpFolderTextBox1, tpFolderDeleteButton1));
            _TPSets.Add(new TPSet(tpFolderCheckBox2, tpFolderButton2, tpFolderTextBox2, tpFolderDeleteButton2));
            _TPSets.Add(new TPSet(tpFolderCheckBox3, tpFolderButton3, tpFolderTextBox3, tpFolderDeleteButton3));
            _TPSets.Add(new TPSet(tpFolderCheckBox4, tpFolderButton4, tpFolderTextBox4, tpFolderDeleteButton4));
            _TPSets.Add(new TPSet(tpFolderCheckBox5, tpFolderButton5, tpFolderTextBox5, tpFolderDeleteButton5));
            _TPSets.Add(new TPSet(tpFolderCheckBox6, tpFolderButton6, tpFolderTextBox6, tpFolderDeleteButton6));
            _TPSets.Add(new TPSet(tpFolderCheckBox7, tpFolderButton7, tpFolderTextBox7, tpFolderDeleteButton7));
            _TPSets.Add(new TPSet(tpFolderCheckBox8, tpFolderButton8, tpFolderTextBox8, tpFolderDeleteButton8));
            _TPSets.Add(new TPSet(tpFolderCheckBox9, tpFolderButton9, tpFolderTextBox9, tpFolderDeleteButton9));
            _TPSets.Add(new TPSet(tpFolderCheckBox10, tpFolderButton10, tpFolderTextBox10, tpFolderDeleteButton10));

            ///もっと賢い方法はないものか
        }

        private void SendImage()
        {
            if (copying) return;

            if (_images.Count == 0) return;

            if (_image != null) _image.Dispose();


            imageNumBox.Text = null;

            loadingImageNameLabel.Text = "copying";

            imageNumBox.Refresh();

            loadingImageNameLabel.Refresh();


            _image = Properties.Resources.copying;

            _matAffine.ZoomFit(pictureBox1, _image, _fitMode);

            RedrawImage();



            copying = true;



            CopyProgressBar.Value = 0;

            CopyProgressBar.Visible = true;

            CopyProgressBar.Minimum = 0;

            CopyProgressBar.Maximum = _images.Count;

            ShowImage nowLoadingImage = _images[nowLoading];



            List<ShowImage> deleteFiles = new List<ShowImage>();

            foreach (ShowImage si in _images)
            {
                bool isCopied = false;
                string fowardingPath;

                for (int i = 0; i < 10; i++)
                {
                    if (si.tpBool[i] & (_TPSets[i]._textBox != null) & (_TPSets[i]._textBox.Text != ""))
                    {
                        try
                        {
                            fowardingPath = System.IO.Path.Combine(_TPSets[i]._textBox.Text, si._fileName);

                            if (!File.Exists(fowardingPath))
                            {
                                FileInfo fi = new System.IO.FileInfo(@si._filePath);
                                fi.MoveTo(@fowardingPath);
                                isCopied = true;
                            }
                        }
                        catch
                        {
                            isCopied = false;
                        }
                    }
                }

                if(si.isChecked)
                {
                    if(isCopied)///表示済みかつコピーした画像の処理
                    {
                        if (_CopyOption3 == 1)
                        {
                            deleteFiles.Add(si);
                        }
                    }
                    else///表示したがコピーされなかった画像の処理
                    {
                        if(_CopyOption1 == 1)
                        {
                            deleteFiles.Add(si);
                        }
                    }
                }
                else///未表示の画像の処理
                {
                    if (_CopyOption2 == 1)
                    {
                        deleteFiles.Add(si);
                    }
                }

                CopyProgressBar.Value++;

                //Application.DoEvents();
            }


            foreach (ShowImage file in deleteFiles)
            {
                _images.Remove(file);


                File.Delete(@file._filePath);
            }

            int loadIndex = 0;

            if (_images.Count == 0)
            {
                _image.Dispose();

                pictureBox1.Image = null;

                imageNumBox.Text = null;

                loadingImageNameLabel.Text = "----";

                imageNumBoxConst.Text = null;

                CreateGraphics(pictureBox1, ref _gPic);

                _matAffine = new System.Drawing.Drawing2D.Matrix();
            }
            else
            {
                try
                {
                    loadIndex = _images.IndexOf(nowLoadingImage);

                    if(loadIndex == -1) loadIndex= 0;
                }
                catch
                {
                    loadIndex = 0;
                }

                Console.WriteLine(nowLoading.ToString() + "," + loadIndex.ToString());

                nowLoading = loadIndex;

                if (_image != null) _image.Dispose();

                try
                {
                    _image = new Bitmap(_images[nowLoading]._filePath);
                }
                catch
                {
                    _image = CreateThumbnail(@_images[nowLoading]._filePath);
                }

                // 画像をピクチャボックスに合わせて表示するアフィン変換行列の計算
                _matAffine.ZoomFit(pictureBox1, _image, _fitMode);

                // ピクチャボックスの背景で画像を削除
                _gPic.Clear(pictureBox1.BackColor);
                // アフィン変換行列に基づいて画像の描画
                _gPic.DrawImage(_image, _matAffine);
                // 更新
                pictureBox1.Refresh();

                imageNumBox.Text = (nowLoading + 1).ToString();
                imageNumBoxConst.Text = "/" + _images.Count.ToString();

                for (int index = 0; index < _TPSets.Count; index++)
                {
                    _TPSets[index]._checkBox.Checked = _images[nowLoading].tpBool[index];

                    if (_TPSets[index]._checkBox.Checked)
                    {
                        _TPSets[index]._textBox.BackColor = SystemColors.GradientActiveCaption;
                    }
                    else
                    {
                        _TPSets[index]._textBox.BackColor = Color.White;
                    }
                }

                loadingImageNameLabel.Text = _images[nowLoading]._fileName;
            }

            ActiveControl = null;

            pictureBox1.Focus();

            CopyProgressBar.Visible = false;

            copying = false;
        }
        private void CopyButton_Click(object sender, EventArgs e)
        {
            if (copying) return;

            SendImage();
        }

        protected override bool ProcessDialogKey(Keys keyData)
        {
            return base.ProcessDialogKey(keyData);
        }

        private void nextButton_Click(object sender, EventArgs e)
        {
            ChangeImage(true);
        }

        private void previousButton_Click(object sender, EventArgs e)
        {
            ChangeImage(false);
        }

        private void MoveToButton_Click(object sender, EventArgs e)
        {
            if (copying) return;


            if (MoveTextbox.Text == null) return;

            int i;

            if (int.TryParse(MoveTextbox.Text, out i))
            {
                i--;

                if (i < 0 | i >= _images.Count) return;

                LoadImage(i);
            }
        }

        private void MoveTextbox_KeyDown(object sender, KeyEventArgs e)
        {
            if(copying) return;

            if(e.KeyCode == Keys.Enter)
            {
                if (MoveTextbox.Text == null) return;

                int i;

                if (int.TryParse(MoveTextbox.Text, out i))
                {
                    i--;

                    if (i < 0 | i >= _images.Count) return;

                    LoadImage(i);
                }
            }
        }


        void LoadImage(int i)
        {
            if (copying) return;

            if (_images.Count == 0) return;

            nowLoading = i;


            if(_image != null) _image.Dispose();

            try
            {
                _image = new Bitmap(_images[nowLoading]._filePath);
            }
            catch
            {
                _image = CreateThumbnail(@_images[nowLoading]._filePath);
            }

            // 画像をピクチャボックスに合わせて表示するアフィン変換行列の計算
            _matAffine.ZoomFit(pictureBox1, _image, _fitMode);

            // 画像の描画
            RedrawImage();

            imageNumBox.Text = (nowLoading + 1).ToString();

            for(int index = 0; index < _TPSets.Count; index++)
            {
                _TPSets[index]._checkBox.Checked = _images[nowLoading].tpBool[index];

                if (_TPSets[index]._checkBox.Checked)
                {
                    _TPSets[index]._textBox.BackColor = SystemColors.GradientActiveCaption;
                }
                else
                {
                    _TPSets[index]._textBox.BackColor = Color.White;
                }
            }

            _images[i].isChecked = true;

            loadingImageNameLabel.Text = _images[nowLoading]._fileName;
        }

        private void pictureBox1_DragDrop(object sender, DragEventArgs e)
        {
            if (copying) return;

            string[] sFileName = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            if (sFileName.Length <= 0)
            {
                return;
            }

            workFolderText.Text = "";

            FileAttributes attr1 = File.GetAttributes(sFileName[0]);

            int i = 0;

            if ((attr1 & FileAttributes.Directory) != FileAttributes.Directory)
            {
                workFolderText.Text = System.IO.Path.GetDirectoryName(sFileName[0]);
                GetImages();

                foreach (ShowImage si in _images)
                {
                    if (si._filePath == sFileName[0]) break;
                    else i++;
                }
            }
            else
            {
                workFolderText.Text = sFileName[0];
                GetImages();
            }

            LoadImage(i);

            ActiveControl = null;

            imageNumBox.Text = (nowLoading + 1).ToString();
            imageNumBoxConst.Text = "/" + _images.Count.ToString();

            pictureBox1.Focus();
        }


        private Bitmap CreateThumbnail(string path)
        {
            FileInfo fi = new FileInfo(path);
            ShellFile shellFile = ShellFile.FromFilePath(path);
            Bitmap bmp = shellFile.Thumbnail.LargeBitmap;
            return bmp;
        }


        /// -----------------------画像移動---------------------------

        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            pictureBox1.Focus();
            _oldPoint.X = e.X;
            _oldPoint.Y = e.Y;

            _mouseDownFlg = true;
        }

        private void pictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            // マウスをクリックしながら移動中のとき、画像の移動
            if (_mouseDownFlg == true)
            {
                // 画像の移動
                _matAffine.Translate(e.X - _oldPoint.X, e.Y - _oldPoint.Y, System.Drawing.Drawing2D.MatrixOrder.Append);
                // 画像の描画
                RedrawImage();

                // ポインタ位置の保持
                _oldPoint.X = e.X;
                _oldPoint.Y = e.Y;
            }


            // アフィン変換行列（画像座標→ピクチャボックス座標）の逆行列(ピクチャボックス座標→画像座標)を求める
            // Invertで元の行列が上書きされるため、Cloneを取ってから逆行列
            var invert = _matAffine.Clone();
            invert.Invert();

            var pf = new PointF[] { e.Location };

            // ピクチャボックス座標を画像座標に変換する
            invert.TransformPoints(pf);
        }

        private void pictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            _mouseDownFlg = false;
        }

        private void RedrawImage()
        {
            if (copying | _image == null) return;

            // ピクチャボックスの背景で画像を削除
            _gPic.Clear(pictureBox1.BackColor);
            // アフィン変換行列に基づいて画像の描画
            _gPic.DrawImage(_image, _matAffine);
            // 更新
            pictureBox1.Refresh();
        }

        private void pictureBox1_MouseWheel(object sender, MouseEventArgs e)
        {
            bool shiftKeyFlg = false;
            if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
            {
                shiftKeyFlg = true;
            }


            if (e.Delta > 0)
            {
                if (shiftKeyFlg)
                {
                    // ポインタの位置周りに回転
                    _matAffine.RotateAt(5f, e.Location, System.Drawing.Drawing2D.MatrixOrder.Append);
                }
                else
                {
                    // 拡大
                    // ポインタの位置周りに拡大
                    _matAffine.ScaleAt(1.5f, e.Location);
                }
            }
            else
            {
                if (shiftKeyFlg)
                {
                    // ポインタの位置周りに回転
                    _matAffine.RotateAt(-5f, e.Location, System.Drawing.Drawing2D.MatrixOrder.Append);
                }
                else
                {
                    // 縮小
                    // ポインタの位置周りに縮小
                    _matAffine.ScaleAt(1.0f / 1.5f, e.Location);
                }
            }
            // 画像の描画
            RedrawImage();
        }

        private static void CreateGraphics(PictureBox pic, ref Graphics g)
        {
            if (g != null)
            {
                g.Dispose();
                g = null;
            }
            if (pic.Image != null)
            {
                pic.Image.Dispose();
                pic.Image = null;
            }

            if ((pic.Width == 0) || (pic.Height == 0))
            {
                return;
            }

            pic.Image = new Bitmap(pic.Width, pic.Height);

            g = Graphics.FromImage(pic.Image);
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
        }

        private void pictureBox1_MouseDoubleClick(object sender, EventArgs e)
        {
            // 画像をピクチャボックスに合わせる
            _matAffine.ZoomFit(pictureBox1, _image, _fitMode);

            // 画像の描画
            RedrawImage();
        }

    }


    public class ShowImage
    {
        public string _filePath;
        public string _fileName;
        public List<bool> tpBool = new List<bool>();
        public bool isChecked = false;

        public ShowImage(string filePass)
        {
            _filePath = filePass;
            _fileName = System.IO.Path.GetFileName(filePass);
            for (int i = 0; i < 10; i++)
            {
                tpBool.Add(false);
            }
        }
    }

    public class TPSet
    {
        public CheckBox _checkBox;
        public Button _button;
        public TextBox _textBox;
        public Button _deleteButton;

        public TPSet(CheckBox checkBox, Button button, TextBox textBox, Button dButton)
        {
            _checkBox = checkBox;
            _button = button;
            _textBox = textBox;
            _deleteButton = dButton;
        }
    }
}
