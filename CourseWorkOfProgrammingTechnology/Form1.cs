using System;
using System.Windows.Forms;
using System.Threading;

namespace CourseWorkOfProgrammingTechnology
{
    public partial class Form1 : Form
    {
        private SystemInputData _systemInputData;

        private Thread _threadForGeneratingReviews;
        private bool _isRun = false;

        public Form1()
        {
            InitializeComponent();
            //Making the window size immutable
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            _systemInputData = new SystemInputData();
        }

        private void SelectingInputDataFile_Click(object sender, EventArgs e)
        {
            _systemInputData.ShowFileDialogForInputDateFile(openFileDialog1);
            pathInputDataFile.Text = _systemInputData.FilePath;
        }
        private void SelectingTeacherSignature_Click(object sender, EventArgs e)
        {
            _systemInputData.ShowFileDialogForSignature(openFileDialog1);
            pathToTeacherSignature.Text = _systemInputData.PathToTeacherSignature;
        }

        private void SelectingFolderToSave_Click(object sender, EventArgs e)
        {
            _systemInputData.ShowFolderDialog(folderBrowserDialog1);
            pathFolderToSave.Text = _systemInputData.FolderPath;
        }

        private void Ok_Click(object sender, EventArgs e)
        {
            if (_systemInputData.IsInputCompleted() && _isRun == false)
            {
                (sender as Button).Enabled = false;
                SelectingInputDataFile.Enabled = false;
                SelectingTeacherSignature.Enabled = false;
                SelectingFolderToSave.Enabled = false;

                _threadForGeneratingReviews = new Thread(() =>
                {
                    var inputData = _systemInputData.GetInputData();
                    GeneratingReviews generatingReviews = new GeneratingReviews(inputData.Item1, inputData.Item2, inputData.Item3, inputData.Item4);
                    generatingReviews.CreateReviews(_systemInputData.FolderPath);

                    (sender as Button).Enabled = true;
                    SelectingInputDataFile.Enabled = true;
                    SelectingTeacherSignature.Enabled = true;
                    SelectingFolderToSave.Enabled = true;

                    _isRun = false;
                });

                _isRun = true;
                _threadForGeneratingReviews.Start();
            }
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


//using System;
//using System.Windows.Forms;
//using System.Threading;

//namespace CourseWorkOfProgrammingTechnology
//{
//    public partial class Form1 : Form
//    {
//        private SystemInputData _systemInputData;

//        public Form1()
//        {
//            InitializeComponent();
//            //Making the window size immutable
//            this.FormBorderStyle = FormBorderStyle.FixedSingle;
//            _systemInputData = new SystemInputData();
//        }

//        private void SelectingInputDataFile_Click(object sender, EventArgs e)
//        {
//            _systemInputData.ShowFileDialogForInputDateFile(openFileDialog1);
//            pathInputDataFile.Text = _systemInputData.FilePath;
//        }
//        private void SelectingTeacherSignature_Click(object sender, EventArgs e)
//        {
//            _systemInputData.ShowFileDialogForSignature(openFileDialog1);
//            pathToTeacherSignature.Text = _systemInputData.PathToTeacherSignature;
//        }

//        private void SelectingFolderToSave_Click(object sender, EventArgs e)
//        {
//            _systemInputData.ShowFolderDialog(folderBrowserDialog1);
//            pathFolderToSave.Text = _systemInputData.FolderPath;
//        }

//        private Thread _threadForGeneratingReviews;
//        private bool _isRun = false;
//        private void Ok_Click(object sender, EventArgs e)
//        {
//            if (_systemInputData.IsInputCompleted() && (_isRun == false))
//            {
//                buttonOk.Enabled = false;
//                SelectingInputDataFile.Enabled = false;
//                SelectingTeacherSignature.Enabled = false;
//                SelectingFolderToSave.Enabled = false;

//                _threadForGeneratingReviews = new Thread(() =>
//                {
//                    var inputData = _systemInputData.GetInputData();
//                    GeneratingReviews generatingReviews = new GeneratingReviews(inputData.Item1, inputData.Item2, inputData.Item3, inputData.Item4);
//                    generatingReviews.CreateReviews(_systemInputData.FolderPath);

//                    buttonOk.Enabled = true;
//                    SelectingInputDataFile.Enabled = true;
//                    SelectingTeacherSignature.Enabled = true;
//                    SelectingFolderToSave.Enabled = true;

//                    _isRun = false;
//                });

//                _isRun = true;
//                _threadForGeneratingReviews.Start();
//            }
//        }

//        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
//        {
//            System.Text.StringBuilder messageBoxCS;
//            if (_threadForGeneratingReviews?.ThreadState == ThreadState.Running)
//            {
//                messageBoxCS = new System.Text.StringBuilder();
//                messageBoxCS.AppendFormat("Программа все еще создает рецензии. Дождитесь завершения");
//                messageBoxCS.AppendLine();
//                MessageBox.Show(messageBoxCS.ToString(), "");

//                e.Cancel = true;
//            }
//        }
//    }
//}
