using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing.Printing;

namespace ISTestTask
{
	public partial class FormMain : Form
	{
		private Education _tempEducation;
		public FormMain()
		{
			InitializeComponent();
			_tempEducation = GetEducation();
		}

		private void buttonSave_Click(object sender, EventArgs e)
		{
			EmployeePersonalCard docData = CreateDocument();

			DocumentsList<EmployeePersonalCard> docsList = new DocumentsList<EmployeePersonalCard>(EmployeePersonalCard.FileName);
			docsList.Deserialize();
			docsList.Add(docData);
			docsList.Serialize();

			//-----------------------------------------//
			Properties.Settings.Default.tabNumberCount++;
			Properties.Settings.Default.docNumberCount++;
			Properties.Settings.Default.Save();
			//-----------------------------------------//
		}

		private EmployeePersonalCard CreateDocument()
		{
			Employee employee = new Employee();

			employee.FamilyName = textBoxFamilyName.Text;
			employee.Name = textBoxName.Text;
			employee.SecondName = textBoxSName.Text;
			employee.DateBirth = textBoxDateBirth.Text;
			employee.PlaceBirth = textBoxPlaceBirth.Text;
			employee.Documets = new EmployeeDocuments(textBoxInn.Text, textBoxSnils.Text);
			employee.KindWork = textBoxKindWork.Text;
			employee.Nationality = textBoxNationality.Text;
			employee.Language = new Language(textBoxLanguage.Text, textBoxLanguageLevel.Text);
			employee.Profession = new Profession(textBoxProfMain.Text, textBoxProfOther.Text);
			employee.TabNumber = Properties.Settings.Default.tabNumberCount;

			if (checkBoxIsTwoEdu.Checked)
			{
				employee.FirstEducation = _tempEducation;
				employee.SecondEducation = GetEducation();
			}
			else
			{
				employee.FirstEducation = GetEducation();
				employee.SecondEducation = new Education();
			}

			employee.ScienceEducation = new Education();
			if (checkBoxScience.Checked)
			{
				employee.ScienceEducation.UniversityName = textBoxScienceName.Text;
				employee.ScienceEducation.DiplomaName = textBoxScienceDoc.Text;
				employee.ScienceEducation.Specialty = textBoxScienceSpecialty.Text;
				employee.ScienceEducation.YearOfEnd = textBoxScienceYearEnd.Text;
				if (radioButtonPostgradute.Checked)
				{
					employee.ScienceEducation.AcademicDegree = radioButtonPostgradute.Text;
				}
				else if (radioButtonGraduate.Checked)
				{
					employee.ScienceEducation.AcademicDegree = radioButtonGraduate.Text;
				}
				else
				{
					employee.ScienceEducation.AcademicDegree = radioButtonDoctorate.Text;
				}
			}

			if (radioButtonTypeWorkMain.Checked) employee.Work = radioButtonTypeWorkMain.Text;
			else employee.Work = radioButtonTypeWorkOther.Text;

			if (radioButtonMale.Checked) employee.Gender = radioButtonMale.Text;
			else employee.Gender = radioButtonFemale.Text;

			EmployeePersonalCard docData = new EmployeePersonalCard(employee);
			docData.DocNumber = Properties.Settings.Default.docNumberCount;

			return docData;
		}

		private Education GetEducation()
		{
			Education education = new Education();
			education.AcademicDegree = textBoxAcademicDegree.Text;
			education.DiplomaName = textBoxDiplomaName.Text;
			education.DiplomaNumber = textBoxDiplomaNumber.Text;
			education.DiplomaSeries = textBoxDiplomaSeries.Text;
			education.Specialty = textBoxSpecialty.Text;
			education.UniversityName = textBoxUniversityName.Text;
			education.YearOfEnd = textBoxYearEnd.Text;

			return education;
		}

		private void comboBoxListDocs_DropDown(object sender, EventArgs e)
		{
			comboBoxListDocs.Items.Clear();

			try
			{
				DocumentsList<EmployeePersonalCard> docsList = new DocumentsList<EmployeePersonalCard>(EmployeePersonalCard.FileName);
				docsList.Deserialize();

				foreach (EmployeePersonalCard doc in docsList)
				{
					comboBoxListDocs.Items.Add(doc.getName());
				}
			}
			catch (ArgumentNullException exp)
			{
				MessageBox.Show(exp.Message);
			}
		}

		private void buttonPrint_Click(object sender, EventArgs e)
		{
			try
			{
				int index = comboBoxListDocs.SelectedIndex;

				DocumentsList<EmployeePersonalCard> list = new DocumentsList<EmployeePersonalCard>(EmployeePersonalCard.FileName);
				list.Deserialize();

				Printer printer = new Printer(list[index]);
			}
			catch (ArgumentOutOfRangeException exp)
			{
				MessageBox.Show("Выберите значение из списка!");
			}

			Word.Application wordApp = new Word.Application();
			wordApp.Visible = false;

			PrintDialog pDialog = new PrintDialog();
			if (pDialog.ShowDialog() == DialogResult.OK)
			{
				string path = Environment.CurrentDirectory;
				Word.Document doc = wordApp.Documents.Add(path + @"\OutputDocument.docx");
				wordApp.ActivePrinter = pDialog.PrinterSettings.PrinterName;
				wordApp.ActiveDocument.PrintOut();
				doc.Close(SaveChanges: false);
				doc = null;
			}
			((Word._Application)wordApp).Quit(SaveChanges: false);

			wordApp = null;
		}

		private void checkBox2_CheckedChanged(object sender, EventArgs e)
		{
			if (checkBoxScience.Checked)
			{
				panelScience.Visible = true;
				panelScienceType.Visible = true;
			}
			else
			{
				panelScience.Visible = false;
				panelScienceType.Visible = false;
			}
		}

		private void checkBoxIsTwoEdu_CheckedChanged(object sender, EventArgs e)
		{

			if (checkBoxIsTwoEdu.Checked)
			{
				_tempEducation = GetEducation();
				textBoxUniversityName.Text = "";
				textBoxSpecialty.Text = "";
				textBoxAcademicDegree.Text = "";
				textBoxDiplomaName.Text = "";
				textBoxDiplomaNumber.Text = "";
				textBoxDiplomaSeries.Text = "";
				textBoxYearEnd.Text = "";
			}
			else
			{
				textBoxUniversityName.Text = _tempEducation.UniversityName;
				textBoxSpecialty.Text = _tempEducation.Specialty;
				textBoxAcademicDegree.Text = _tempEducation.AcademicDegree;
				textBoxDiplomaName.Text = _tempEducation.DiplomaName;
				textBoxDiplomaNumber.Text = _tempEducation.DiplomaNumber;
				textBoxDiplomaSeries.Text = _tempEducation.DiplomaSeries;
				textBoxYearEnd.Text = _tempEducation.YearOfEnd;
			}
		}

		void ResetTextBoxes(System.Windows.Forms.Control.ControlCollection controls)
		{
			foreach (Control c in controls)
			{
				TextBox tb = c as TextBox;
				if (tb != null)
				{
					tb.Text = string.Empty;
				}
				ResetTextBoxes(c.Controls);
			}
		}

		private void buttonCreate_Click(object sender, EventArgs e)
		{
			ResetTextBoxes(this.Controls);
		}

		private void MenuItemCreate_Click(object sender, EventArgs e)
		{
			buttonCreate.PerformClick();
		}

		private void MenuItemSave_Click(object sender, EventArgs e)
		{
			buttonSave.PerformClick();
		}

		private void MenuItemExit_Click(object sender, EventArgs e)
		{
			Application.Exit();
		}
	}
}