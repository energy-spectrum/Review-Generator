using System;
using System.Windows.Forms;
using System.Threading;

namespace CourseWorkOfProgrammingTechnology
{
    public partial class Form1 : Form
    {
        private string _pathToExcelFile = string.Empty;
        private string _pathToTeacherSignature = string.Empty; 
        private string _pathToFolderToSave = string.Empty;

        private InitialDataReader _initialDataReader;

        private Thread _threadForGeneratingReviews;
        private bool _isRun = false;

        public Form1()
        {
            InitializeComponent();
            //Making the window size immutable
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            this._initialDataReader = new InitialDataReader();
        }

        private void SelectingInputDataFile_Click(object sender, EventArgs e)
        {
            _pathToExcelFile = ShowFileDialog(openFileDialog1, filter: "excel files (*.xlsx)|*.xlsx");
            pathInputDataFile.Text = _pathToExcelFile;
        }
        private void SelectingTeacherSignature_Click(object sender, EventArgs e)
        {
            _pathToTeacherSignature = ShowFileDialog(openFileDialog1, filter: "(*.jp2)|*.jp2| (*.jpg)|*.jpg| (*.png)|*.png");
            pathToTeacherSignature.Text = _pathToTeacherSignature;
        }
        private string ShowFileDialog(OpenFileDialog openFileDialog, string filter)
        {
            string universalPathToDesktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            openFileDialog.InitialDirectory = universalPathToDesktop;
            openFileDialog.Filter = filter;
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            string filePath = string.Empty;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
            }

            return filePath;
        }

        private void SelectingFolderToSave_Click(object sender, EventArgs e)
        {
            _pathToFolderToSave = string.Empty;

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                _pathToFolderToSave = folderBrowserDialog1.SelectedPath;
            }
            pathFolderToSave.Text = _pathToFolderToSave;
        }

        public bool IsInputCompleted()
        {
            return (_pathToExcelFile != string.Empty) && (_pathToTeacherSignature != string.Empty) && (_pathToFolderToSave != string.Empty);
        }

        private void Ok_Click(object sender, EventArgs e)
        {
            if (IsInputCompleted() && _isRun == false)
            {
                DisableButtons();

                _threadForGeneratingReviews = new Thread(() =>
                {
                    InitialData initialData = _initialDataReader.GetInitialDataFromFile(_pathToExcelFile);

                    ReviewsGenerator reviewsGenerator = new ReviewsGenerator(initialData, _pathToTeacherSignature, _pathToFolderToSave);
                    reviewsGenerator.CreateReviews();

                    EnableButtons();
                    _isRun = false;
                });

                _isRun = true;
                _threadForGeneratingReviews.Start();
            }
        }
        private void DisableButtons()
        {
            ToggleButtons(enabled: false);
        }

        //Возникает исключение: System.InvalidOperationException: 
        //"Недопустимая операция в нескольких потоках: попытка доступа к элементу управления 'buttonOk' не из того потока,
        //в котором он был создан."
        //Аналогично и к элементам управления SelectingInputDataFile, SelectingTeacherSignature, SelectingFolderToSave
        //Поэтому включение кнопок происходит в цикле while. До тех пор, пока последняя кнопка не будет включена
        private void EnableButtons()
        {
            while (SelectingFolderToSave.Enabled == false)
            {
                try
                {
                    ToggleButtons(enabled: true);
                }
                catch { }
            }
        }

        private void ToggleButtons(bool enabled)
        {
            buttonOk.Enabled = enabled;
            SelectingInputDataFile.Enabled = enabled;
            SelectingTeacherSignature.Enabled = enabled;
            SelectingFolderToSave.Enabled = enabled;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            System.Text.StringBuilder messageBoxCS;
            if (_threadForGeneratingReviews?.ThreadState == ThreadState.Running)
            {
                messageBoxCS = new System.Text.StringBuilder();
                messageBoxCS.AppendFormat("Программа все еще создает рецензии. Дождитесь завершения");
                messageBoxCS.AppendLine();
                MessageBox.Show(messageBoxCS.ToString(), "");

                e.Cancel = true;
            }
        }
    }
}