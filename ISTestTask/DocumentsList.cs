using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

namespace ISTestTask
{
	public class DocumentsList<T>
	{
		private List<T> List = new List<T>();
		private string _fileName = "documents.xml";

		public DocumentsList()
		{

		}

		public DocumentsList(string fileName)
		{
			_fileName = fileName;
		}

		public void Add(T o)
		{
			List.Add(o);
		}

		public int Count()
		{
			return List.Count;
		}

		public void Serialize()
		{
			XmlSerializer serializer = new XmlSerializer(typeof(List<T>), new XmlRootAttribute("DocumentsList"));

			XmlWriterSettings settings = new XmlWriterSettings() { Indent = true };

			using (XmlWriter writer = XmlTextWriter.Create(_fileName, settings))
			{
				serializer.Serialize(writer, List);
			}
		}

		public void Deserialize()
		{
			XmlSerializer serializer = new XmlSerializer(typeof(List<T>), new XmlRootAttribute("DocumentsList"));

			if (File.Exists(_fileName))
			{
				using (XmlReader reader = XmlTextReader.Create(_fileName))
				{
					try
					{
						List = (List<T>)serializer.Deserialize(reader);
					}
					catch (InvalidOperationException e)
					{
						MessageBox.Show(e.Message);
						throw new InvalidOperationException();
					}
				}
			}
		}

		public IEnumerator<T> GetEnumerator()
		{
			foreach (T doc in List)
			{
				yield return doc;
			}
		}

		public T this[int index]
		{
			get
			{
				return List[index];
			}
			set
			{
				List[index] = value;
			}
		}

		public string FileName
		{
			get
			{
				return _fileName;
			}
			set
			{
				_fileName = value;
			}
		}
	}
}
