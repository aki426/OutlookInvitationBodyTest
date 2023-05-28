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
using ConsoleApp1.AppointmentItem.Tests;

namespace ConsoleApp1.Utility.Tests
{
	[TestClass()]
	public class IndexedStringBuilderTests
	{
		private static Outlook.Application application;

		[ClassInitialize]
		public static void ClassInitialize(TestContext testContext)
		{
			// Outlookインスタンスはクラス内で使いまわす。
			application = new Outlook.Application();
		}

		[TestMethod()]
		public void IndexedLineBufferTest()
		{
			List<(int index, char c)> expected = new List<(int index, char c)>();

			IndexedStringBuilder buf = new IndexedStringBuilder();
			(int start, int end) sec;
			_ = buf.AppendLine();
			_ = buf.AppendParagraph();
			sec = buf.Append("ほげ");
			expected.Add((sec.start, 'ほ'));
			expected.Add((sec.end, 'げ'));
			sec = buf.AppendLine("ぴよ  ");
			expected.Add((sec.start, 'ぴ'));
			expected.Add((sec.end, ' '));
			sec = buf.AppendParagraph("ふがもげ");
			expected.Add((sec.start, 'ふ'));
			expected.Add((sec.end, 'げ'));

			// AppointmentItemの作成。
			Outlook.AppointmentItem appointmentItem =
				application.CreateItem(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;

			// 元の文字列。
			string str = buf.ToString();
			AppointmentItemControllerTests.PrintChars(str);
			Console.WriteLine("");
			Console.WriteLine("----------------------------------------");

			appointmentItem.Body = str; // "hogehoge";

			// Body
			AppointmentItemControllerTests.PrintChars(appointmentItem.Body);
			Console.WriteLine("");
			Console.WriteLine("----------------------------------------");

			appointmentItem.Display(); // 表示

			// インビテーションはバックカラー、文字色、文字サイズを変更できないため
			// インビテーションを展開した後にWordオブジェクトを使用して色付けと文字サイズの変更を行う
			// WordのDocumentオブジェクト経由で書式設定を行う。
			string actual = string.Empty;
			Outlook.Inspector ins = application.ActiveInspector();
			if (ins.EditorType == Outlook.OlEditorType.olEditorWord)
			{
				Word.Document doc = ins.WordEditor as Word.Document;

				// Content.Text
				AppointmentItemControllerTests.PrintChars(doc.Content.Text);
				Console.WriteLine("");

				actual = doc.Content.Text;
			}

            foreach ((int index, char c) ex in expected)
            {
				Console.WriteLine($"[{ex.index}] - '{ex.c}'");
                if (ex.c == ' ' && (actual[ex.index] == '\u00a0' || actual[ex.index] == ' '))
                {
					// NOTE: [SP]は[Nbsp]へ変換されてしまうのでどちらでもマッチするようにする。
					// nop
				}
				else
				{
					Assert.AreEqual(ex.c, actual[ex.index]);
				}
            }

            appointmentItem.Close(Outlook.OlInspectorClose.olDiscard);
		}
	}
}
