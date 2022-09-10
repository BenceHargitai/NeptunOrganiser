using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.IO;
using SautinSoft.Document;
using System.Globalization;
using System.Windows.Forms;


using HorizontalAlignment = SautinSoft.Document.HorizontalAlignment;
using System.Diagnostics;

namespace NeptunOrganiser
{
	public partial class Form1 : Form
	{
		class Tanora
		{
			public DateTime start { get; set; }
			public DateTime end { get; set; }
			public string description { get; set; }
			public Tanora()
			{
				start = new DateTime();
				end = new DateTime();
				description = "";
			}
		}
		public static string nevgeneralas()
		{
			string ev = DateTime.Now.Year.ToString() + ".";
			string honap = DateTime.Now.Month.ToString()[0].ToString() != "0" ? "0" + DateTime.Now.Month.ToString()+"." : DateTime.Now.Month.ToString() + ".";
			string nap =  DateTime.Now.Day.ToString()[0].ToString() != "0" ? "0" + DateTime.Now.Day.ToString() +".": DateTime.Now.Day.ToString() + ".";

			return ev + honap + nap;
		}
		public static void CreateDocxUsingDocumentBuilder()
		{
			// Set a path to our document.
			DocumentCore dc = new DocumentCore();
			string nev = "ProgramozasAlapok20 - " + nevgeneralas();
			string filePath = @"A:\" + nev + ".docx";

			ParagraphStyle paragraphStyle1 = new ParagraphStyle("ParagraphStyle1");
			paragraphStyle1.CharacterFormat.FontName = "Avenir Next LT Pro Demi";
			paragraphStyle1.CharacterFormat.Size = 9;
			dc.Styles.Add(paragraphStyle1);

			ParagraphStyle paragraphStyle2 = new ParagraphStyle("ParagraphStyle2");
			paragraphStyle2.CharacterFormat.FontName = "Avenir Next LT Pro Demi";
			paragraphStyle2.CharacterFormat.Size = 22;
			dc.Styles.Add(paragraphStyle2);

			ParagraphStyle paragraphStyle3 = new ParagraphStyle("ParagraphStyle3");
			paragraphStyle3.CharacterFormat.FontName = "Bahnschrift SemiBold";
			paragraphStyle3.CharacterFormat.Size = 11;
			dc.Styles.Add(paragraphStyle3);


			dc.Sections.Add(
				new Section(dc,
					new Paragraph(dc, (DateTime.Now.Year).ToString() + "." + (DateTime.Now.Month).ToString() + "." + (DateTime.Now.Day).ToString())
					{
						ParagraphFormat = new ParagraphFormat
						{
							Alignment = HorizontalAlignment.Right,
							Style = paragraphStyle1

						}
					},
					new Paragraph(dc, "Programozás alapok 20.")
					{
						ParagraphFormat = new ParagraphFormat
						{
							LineSpacingRule = LineSpacingRule.Exactly,
							SpaceBefore = 0,
							SpaceAfter = 34,
							Alignment = HorizontalAlignment.Center,
							Style = paragraphStyle2
						}
					},
					new Paragraph(dc, "")
					{
						ParagraphFormat = new ParagraphFormat
						{
							LineSpacingRule = LineSpacingRule.Exactly,
							SpaceBefore = 10,
							Alignment = HorizontalAlignment.Left,
							Style = paragraphStyle3
						}
					}));

			dc.Save(filePath);
			System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
		}
		public static void SLNProjektManager()
		{

			string targetPath = @"A:\";
			bool useCurrentTime = true;
			byte mode = 0;
			using (StreamReader f = new StreamReader("settings.txt"))
			{
				string[] temp = f.ReadLine().Split('"');
				//targetPath = temp[1];
				Console.WriteLine(temp[1]);
				temp = f.ReadLine().Split('"');
				if (temp[1] == "False")
					useCurrentTime = false;

				temp = f.ReadLine().Split('"');
				if (temp[1] == "1")
					mode = 1;
			}

			if (useCurrentTime)
				targetPath += @"\" + DateTime.Now.ToString("yyyy.MM.dd") + " - " + DateTime.Now.ToString("HH") + "h" + DateTime.Now.ToString("mm") + "m";

			List<string> sourcePath = new List<string>();
			sourcePath.Add(Directory.GetCurrentDirectory() + @"\csharp_template");
			sourcePath.Add(Directory.GetCurrentDirectory() + @"\cplusplus_template");

			Console.WriteLine(Directory.GetCurrentDirectory() + @"\csharp_template");

			foreach (string dirPath in Directory.GetDirectories(sourcePath[mode], "*", SearchOption.AllDirectories))
			{
				Directory.CreateDirectory(dirPath.Replace(sourcePath[mode], targetPath));
			}

			//Copy all the files & Replaces any files with the same name
			foreach (string newPath in Directory.GetFiles(sourcePath[mode], "*.*", SearchOption.AllDirectories))
			{
				File.Copy(newPath, newPath.Replace(sourcePath[mode], targetPath), true);
			}
		}
		public static void Beolvaso()
		{
			List<Tanora> tanorak = new List<Tanora>();

			try
			{
				using (StreamReader f = new StreamReader("CalendarExport.ics"))
					while (!f.EndOfStream)
					{
						string[] temp = f.ReadLine().Split(':');
						if (temp[0] == "DTSTART")
						{
							temp[1] = temp[1].Remove(8, 1);
							temp[1] = temp[1].Remove(14, 1);

							tanorak.Add(new Tanora());
							tanorak[tanorak.Count - 1].start = DateTime.ParseExact(temp[1], "yyyyMMddHHmmss", CultureInfo.InvariantCulture);
						}
						else if (temp[0] == "DTEND")
						{
							temp[1] = temp[1].Remove(8, 1);
							temp[1] = temp[1].Remove(14, 1);

							tanorak[tanorak.Count - 1].end = DateTime.ParseExact(temp[1], "yyyyMMddHHmmss", CultureInfo.InvariantCulture);
						}
						else if (temp[0] == "SUMMARY")
							tanorak[tanorak.Count - 1].description = temp[1];
					}
			}
			catch(FileNotFoundException)
			{
				MessageBox.Show("Az adott fájl nem található");
			}
		}


		public Form1()
		{
			InitializeComponent();
			topPanel.BackColor = ColorTranslator.FromHtml("#2995fe");
			panel2.Cursor = System.Windows.Forms.Cursors.Hand;
			panel3.Cursor = System.Windows.Forms.Cursors.Hand;
			panel6.Cursor = System.Windows.Forms.Cursors.Hand;

			if (System.IO.File.Exists("CalendarExport.ics"))
			{
				label10.Hide();
			}

		}

		private void Form1_Load(object sender, EventArgs e)
		{
			nevgeneralas();
		}



		private void DocxGeneralas(object sender, EventArgs e)
		{
			CreateDocxUsingDocumentBuilder();

		}

		private void CsharpProjekt(object sender, EventArgs e)
		{
			SLNProjektManager();
		}


		private void CppProjekt(object sender, EventArgs e)
		{
			SLNProjektManager();
		}


		private void button1_Click_1(object sender, EventArgs e)
		{

			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.CheckFileExists = true;
			openFileDialog.AddExtension = true;
			openFileDialog.Multiselect = false;
			openFileDialog.Filter = "Naptár fileok|*.ics";

			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				var fileName = openFileDialog.FileName;
				try
				{
					System.IO.File.Copy(fileName, Path.Combine(@".\", "CalendarExport" + ".ics"));
					MessageBox.Show("Az importálás sikeres volt!");
					label10.Hide();
				}
				catch
				{
					DialogResult dialogResult = MessageBox.Show("Már van beimportálva egy órarend. Kívánod felülírni?", "Import", MessageBoxButtons.YesNo);
					if (dialogResult == DialogResult.Yes)
					{
						System.IO.File.Copy(fileName, Path.Combine(@".\", "CalendarExport" + ".ics"), true);
						MessageBox.Show("Az importálás sikeres volt!");

					}
					else
					{
						MessageBox.Show("Az importálás sikertelen volt!");
					}
				}
			}


		}
	}
}
