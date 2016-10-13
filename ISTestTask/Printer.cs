using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using TemplateEngine.Docx;
using System.Windows.Forms;
using System.Reflection;

namespace ISTestTask
{
	class Printer
	{
		private EmployeePersonalCard _employeePersonalCard;

		public Printer(EmployeePersonalCard employeePersonalCard)
		{
			this._employeePersonalCard = employeePersonalCard;
			InitializePrint();
		}

		private void InitializePrint()
		{
			Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ISTestTask.Resources.InputTemplate.docx");
			var fileName = Path.GetTempFileName();

			try
			{
				using (FileStream fs = File.OpenWrite(fileName))
				{
					stream.CopyTo(fs);
				}
				File.Delete("OutputDocument.docx");
				File.Copy(fileName, "OutputDocument.docx");

				var valuesToFill = new Content(
						new FieldContent("Report date", DateTime.Now.ToString("dd:MM:yyyy")),
						new FieldContent("Tab number", _employeePersonalCard.Employee.TabNumber.ToString()),
						new FieldContent("Inn", _employeePersonalCard.Employee.Documets.Inn),
						new FieldContent("Snils", _employeePersonalCard.Employee.Documets.Snils),
						new FieldContent("Kind work", _employeePersonalCard.Employee.KindWork),
						new FieldContent("Type work", _employeePersonalCard.Employee.Work),
						new FieldContent("Gender", _employeePersonalCard.Employee.Gender),
						new FieldContent("Doc number", _employeePersonalCard.DocNumber.ToString()),
						new FieldContent("Family name", _employeePersonalCard.Employee.FamilyName),
						new FieldContent("Name", _employeePersonalCard.Employee.Name),
						new FieldContent("Second name", _employeePersonalCard.Employee.SecondName),
						new FieldContent("Date birth", _employeePersonalCard.Employee.DateBirth),
						new FieldContent("Place birth", _employeePersonalCard.Employee.PlaceBirth),
						new FieldContent("Profession main", _employeePersonalCard.Employee.Profession.Main),
						new FieldContent("Profession other", _employeePersonalCard.Employee.Profession.Other),
						new FieldContent("Language", _employeePersonalCard.Employee.Language.Name),
						new FieldContent("Language level", _employeePersonalCard.Employee.Language.Level),
						new FieldContent("Nationality", _employeePersonalCard.Employee.Nationality),
						new FieldContent("First degree", _employeePersonalCard.Employee.FirstEducation.AcademicDegree),
						new FieldContent("First diploma name", _employeePersonalCard.Employee.FirstEducation.DiplomaName),
						new FieldContent("First diploma number", _employeePersonalCard.Employee.FirstEducation.DiplomaNumber),
						new FieldContent("First diploma series", _employeePersonalCard.Employee.FirstEducation.DiplomaSeries),
						new FieldContent("First specialty", _employeePersonalCard.Employee.FirstEducation.Specialty),
						new FieldContent("First university", _employeePersonalCard.Employee.FirstEducation.UniversityName),
						new FieldContent("First year end", _employeePersonalCard.Employee.FirstEducation.YearOfEnd),
						new FieldContent("Second degree", _employeePersonalCard.Employee.SecondEducation.AcademicDegree),
						new FieldContent("Second diploma name", _employeePersonalCard.Employee.SecondEducation.DiplomaName),
						new FieldContent("Second diploma number", _employeePersonalCard.Employee.SecondEducation.DiplomaNumber),
						new FieldContent("Second diploma series", _employeePersonalCard.Employee.SecondEducation.DiplomaSeries),
						new FieldContent("Second specialty", _employeePersonalCard.Employee.SecondEducation.Specialty),
						new FieldContent("Second university", _employeePersonalCard.Employee.SecondEducation.UniversityName),
						new FieldContent("Second year end", _employeePersonalCard.Employee.SecondEducation.YearOfEnd),
						new FieldContent("Science name", _employeePersonalCard.Employee.ScienceEducation.AcademicDegree),
						new FieldContent("Science university", _employeePersonalCard.Employee.ScienceEducation.UniversityName),
						new FieldContent("Science diploma", _employeePersonalCard.Employee.ScienceEducation.DiplomaName),
						new FieldContent("Science specialty", _employeePersonalCard.Employee.ScienceEducation.Specialty),
						new FieldContent("Science year end", _employeePersonalCard.Employee.ScienceEducation.YearOfEnd)
						);

				using (var outputDocument = new TemplateProcessor("OutputDocument.docx").SetRemoveContentControls(true))
				{
					outputDocument.FillContent(valuesToFill);
					outputDocument.SaveChanges();
				}
			}
			catch (IOException e)
			{
				MessageBox.Show(e.Message);
			}
			finally
			{
				File.Delete(fileName);
			}
		}
	}
}
