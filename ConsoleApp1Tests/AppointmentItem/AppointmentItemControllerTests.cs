using Microsoft.VisualStudio.TestTools.UnitTesting;
using ConsoleApp1.AppointmentItem;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using ConsoleApp1.Utility;

namespace ConsoleApp1.AppointmentItem.Tests
{
	[TestClass()]
	public class AppointmentItemControllerTests
	{
		private static Outlook.Application application;

		[ClassInitialize]
		public static void ClassInitialize(TestContext testContext)
		{
			// Outlookインスタンスはクラス内で使いまわす。
			application = new Outlook.Application();
		}

		[DataTestMethod]
		[DataRow("001", "")]
		[DataRow("002", " ")]
		[DataRow("003", "  ")]
		[DataRow("004", "   ")]
		[DataRow("004.5", "    ")]
		[DataRow("005", "ほげ")]
		[DataRow("006", " ほげ")]
		[DataRow("007", "  ほげ")]
		[DataRow("008", "   ほげ")]
		[DataRow("008.5", "    ほげ")]
		[DataRow("009", "ほげ ")]
		[DataRow("010", "ほげ  ")]
		[DataRow("011", "ほげ   ")]
		[DataRow("012", "ほ げ")]
		[DataRow("013", "ほ  げ")]
		[DataRow("014", "ほ   げ")]
		[DataRow("015", @"ほげ
ぴよ
ふが")]
		[DataRow("016", @"ほげ
ぴよ

ふが


もげ



ほげぴよ")]
		[DataRow("017", " ほ げ ")] // 半角スペースはSP
		[DataRow("018", " ほ げ ")] // 半角スペースは'\u00a0'
		[DataRow("019", "  ほ  げ  ")] // 半角スペースは'\u00a0'
		[DataRow("020", "   ほ   げ   ")] // 半角スペースは'\u00a0'
		[DataRow("021", "    ほ    げ    ")] // 半角スペースは'\u00a0'
		[DataRow("022", "　ほ　げ　")] // 全角スペース
		[DataRow("023", "　　ほ　　げ　　")] // 全角スペース
		[DataRow("024", " ")] // 半角スペースは'\u00a0'
		[DataRow("025", @"
")]
		[DataRow("026", @"

")]
		[DataRow("027", @"


")]
		[DataRow("028", @"
ほげ")]
		[DataRow("029", @"

ほげ")]
		[DataRow("030", @"


ほげ")]
		[DataRow("031", @" ほげ ぴよ 
ふが")]
		[DataRow("032", @"  ほげ  ぴよ  
ふが")]
		[DataRow("033", @"   ほげ   ぴよ   
ふが")]
		[DataRow("034", @"ほげ
ぴよ
 
ふが
 
 
もげ")]
		[DataRow("035", @"
ほげ")]
		[DataRow("036", @"

ほげ")]
		[DataRow("037", @"


ほげ")]
		[DataRow("038", @" 
 
ほげ")]
		[DataRow("039", "ほ\r\nげ")]
		[DataRow("040", "ほ \r\n \r\nげ")]
		[DataRow("041", "ほ  \r\n  \r\nげ")]
		[DataRow("042", "ほ   \r\n   \r\nげ")]


		public void CreateApointmentItemTest(string nouse, string str)
		{

			// AppointmentItemの作成。
			Outlook.AppointmentItem appointmentItem =
				application.CreateItem(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;

			// 元の文字列。
			PrintChars(str);
			Console.WriteLine("");
			Console.WriteLine("----------------------------------------");

			appointmentItem.Body = str; // "hogehoge";

			// Body
			PrintChars(appointmentItem.Body);
			Console.WriteLine("");
			Console.WriteLine("----------------------------------------");

			appointmentItem.Display(); // 表示

			// インビテーションはバックカラー、文字色、文字サイズを変更できないため
			// インビテーションを展開した後にWordオブジェクトを使用して色付けと文字サイズの変更を行う
			// WordのDocumentオブジェクト経由で書式設定を行う。
			Outlook.Inspector ins = application.ActiveInspector();
			if (ins.EditorType == Outlook.OlEditorType.olEditorWord)
			{
				Word.Document doc = ins.WordEditor as Word.Document;

				// Content.Text
				PrintChars(doc.Content.Text);
				Console.WriteLine("");
			}

			appointmentItem.Close(Outlook.OlInspectorClose.olDiscard);
		}

		// \r \n \u00a0 \v

		[DataTestMethod]
		[DataRow("001", "")]
		[DataRow("002", " ")]
		[DataRow("003", " \r\n  \r\n   \r\n    \r\nZ")]
		[DataRow("004", " \r\n\r\n  \r\n\r\n   \r\n\r\n    \r\n\r\nZ")]
		[DataRow("005", "a\r\nあ \r\nい  \r\nう   \r\nえ    \r\nZ")]
		[DataRow("006", "a\r\n\r\nあ \r\n\r\nい  \r\n\r\nう   \r\n\r\nえ    \r\n\r\nZ")]
		public void WordContentTextTest(string nouse, string str)
		{

			// AppointmentItemの作成。
			Outlook.AppointmentItem appointmentItem =
				application.CreateItem(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;

			// 元の文字列。
			PrintChars(str);
			Console.WriteLine("");
			Console.WriteLine("----------------------------------------");

			appointmentItem.Body = str; // "hogehoge";

			// Body
			PrintChars(appointmentItem.Body);
			Console.WriteLine("");
			Console.WriteLine("----------------------------------------");

			appointmentItem.Display(); // 表示

			// インビテーションはバックカラー、文字色、文字サイズを変更できないため
			// インビテーションを展開した後にWordオブジェクトを使用して色付けと文字サイズの変更を行う
			// WordのDocumentオブジェクト経由で書式設定を行う。
			Outlook.Inspector ins = application.ActiveInspector();
			if (ins.EditorType == Outlook.OlEditorType.olEditorWord)
			{
				Word.Document doc = ins.WordEditor as Word.Document;

				// Content.Text
				PrintChars(doc.Content.Text);
				Console.WriteLine("");
			}

			appointmentItem.Close(Outlook.OlInspectorClose.olDiscard);
		}

		/// <summary>
		/// StringをCharに分解してユニコード付きで表示する便利関数。
		/// </summary>
		/// <param name="str"></param>
		public static void PrintChars(string str)
		{
			Console.WriteLine($"\"{str}\"");

			if (str == string.Empty)
			{
				Console.WriteLine("string.Empty.");
			}
			else if (str == null)
			{
				Console.WriteLine("null string.");
			}
			else
			{
				Console.WriteLine($"Length : {str.Length}");

				for (int i = 0; i < str.Length; i++)
				{
					int c = (int)str[i];
					if (c == 10)
					{
						Console.WriteLine($"s[{i:d2}] = <LF> ('\\u{(int)str[i]:x4}')");
					}
					else if (c == 11)
					{
						Console.WriteLine($"s[{i:d2}] = |VT| ('\\u{(int)str[i]:x4}')");
					}
					else if (c == 13)
					{
						Console.WriteLine($"s[{i:d2}] = <CR> ('\\u{(int)str[i]:x4}')");
					}
					else if (c == 32)
					{
						Console.WriteLine($"s[{i:d2}] = [SP] ('\\u{(int)str[i]:x4}')");
					}
					else if (c <= 31)
					{
						Console.WriteLine($"s[{i:d2}] = ___ ('\\u{(int)str[i]:x4}')");
					}
					else // if (0 <= c && c <= 127)
					{
						Console.WriteLine($"s[{i:d2}] = '{str[i]}' ('\\u{(int)str[i]:x4}')");
					}
					//else
					//{
					//	Console.WriteLine($"s[{i:d2}] = '{str[i]}'");
					//}
				}
			}

			Console.WriteLine();
		}

		[DataTestMethod]
		[DataRow("001", " ほ げ ")]
		[DataRow("002", " ")]
		[DataRow("003", "  ")]
		[DataRow("004", "   ")]
		[DataRow("004.5", "    ")]
		public void StringReplaceTest(string nouse, string str)
		{
			// 元の文字列。
			PrintChars(str);
			Console.WriteLine("");
			Console.WriteLine("----------------------------------------");

			// 半角スペース'\u0020'から'\u00a0'
			//PrintChars(str.Replace(" ", " "));
			PrintChars(IndexedLineBuffer.ConvertSpToNbsp(str));
		}
	}
}



