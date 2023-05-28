using System.Text;

namespace ConsoleApp1.Utility
{
	/// <summary>
	/// AppointmentItem.Body経由でWord.Document.Content.Text化される文字列について、
	/// 追加した文字列の先頭からのインデクスを取得することができるStringBuilderのラッパークラス。
	/// 通常のStringからWord.Document.Content.Textへの変換規則については各関数を参照。
	/// </summary>
	public class IndexedStringBuilder
	{
		//////////////////////////////// メンバー変数 ////////////////////////////////

		/// <summary>
		/// 文章全体の文字列バッファ。
		/// </summary>
		private readonly StringBuilder _sentenceBuilder;
		/// <summary>
		/// 現在行の文字列バッファ。
		/// </summary>
		private readonly StringBuilder _lineBuilder;
		/// <summary>
		/// 現在IndexedStringBuilderが蓄えている文字列のWord.Document.Content.Text化した場合の終端インデクス値。
		/// 改行やスペースがStringBuilder内の数ではなく、Word.Document.Content.Textでのそれとしてカウントされていることに注意。
		/// </summary>
		private int _currentIndex;

		//////////////////////////////// Static関数 ////////////////////////////////

		/// <summary>
		/// SpaceをNbspへ変換する関数。
		/// </summary>
		/// <param name="str"></param>
		/// <returns></returns>
		public static string ConvertSpToNbsp(string str)
		{
			return str.Replace("\u0020", "\u00a0");
		}

		/// <summary>
		/// 改行を消去する関数。
		/// </summary>
		/// <param name="str"></param>
		/// <returns></returns>
		public static string DeleteCrAndLf(string str)
		{
			return str.Replace("\u000a", "").Replace("\u000d", "");
		}

		private static readonly char[] SpaceCharArray = { ' ' };

		/// <summary>
		/// 文末のスペースをカウントする関数。
		/// </summary>
		/// <param name="str"></param>
		/// <returns></returns>
		public static int CountEndSpace(string str)
		{
			return str.Length - str.TrimEnd(SpaceCharArray).Length;
		}

		/// <summary>
		/// スペースで構成された文字列であればTrueを返す関数。
		/// NOTE: string.Emptyの場合もTrueを返すので、空文字列の判定は別途実施すること。
		/// </summary>
		/// <param name="str"></param>
		/// <returns></returns>
		public static bool IsAllSpace(string str)
		{
			return str.TrimEnd(SpaceCharArray).Length == 0;
		}
		
		/// <summary>
		/// ある1行分の文字列として見た場合、どのような文字列で構成された行として分類されるかを表すEnum。
		/// </summary>
		public enum LineType
		{
			/// <summary>
			/// 空。
			/// </summary>
			Empty,
			/// <summary>
			/// スペースだけで構成された文字列。
			/// </summary>
			OnlySpaces,
			/// <summary>
			/// スペースで終わるが通常の文字を含む文字列。
			/// </summary>
			EndWithSpaces,
			/// <summary>
			/// スペースで終わらない通常の文字を含む文字列。
			/// </summary>
			EndWithNoSpace
		}

		/// <summary>
		/// ある1行分の文字列として見た場合、どのような文字列で構成された行として分類されるかを返す関数。
		/// </summary>
		/// <param name="str"></param>
		/// <returns></returns>
		public static LineType GetLineType(string str)
		{
			if (str == string.Empty)
			{
				return LineType.Empty;
			}
			else if (IsAllSpace(str))
			{
				return LineType.OnlySpaces;
			}
			else if (CountEndSpace(str) == 0)
			{
				return LineType.EndWithNoSpace;
			}
			else
			{
				return LineType.EndWithSpaces;
			}
		}

		//////////////////////////////// コンストラクタ ////////////////////////////////
		
		/// <summary>
		/// コンストラクタ。改行のインデクスサイズ＝何文字で表現されるべきかを引数として与える。
		/// </summary>
		/// <param name="newLineSize"></param>
		public IndexedStringBuilder()
		{
			this._sentenceBuilder = new StringBuilder();
			this._lineBuilder = new StringBuilder();
			this._currentIndex = 0;
		}

		//////////////////////////////// メンバー関数 ////////////////////////////////

		/// <summary>
		/// 改行せずに文字列を詰め込む。
		/// 引数の文字列中に改行コードCRやLFが含まれると意図しない改行になってしまうため、改行コードは除去される。
		/// </summary>
		/// <param name="str"></param>
		/// <returns>与えられた文字列の開始終了位置のインデクス。</returns>
		public (int start, int end) Append(string str)
		{
			// 改行コード除去
			string appendee = DeleteCrAndLf(str);
			if (appendee == string.Empty)
			{
				// 詰め込むものが無い場合は何もしない。
				return (_currentIndex, _currentIndex);
			}

			// 何かしら詰め込むものがある場合、lineに追加する。
			this._lineBuilder.Append(appendee);

			// Update Index.
			int start = _currentIndex;
			_currentIndex += appendee.Length;
			int end = _currentIndex;

			// Indexは0始まりなのでendは1Index分引く必要がある。
			return (start, end - 1);
		}

		/// <summary>
		/// 文字列を詰め込み、末尾で改行を行う。
		/// Word.Document.Content.Textへ変換した時に[SP][VT]で改行が表現されるように調整する。
		/// </summary>
		/// <param name="str">空文字の場合は空白1文字が追加される。</param>
		/// <returns>与えられた文字列の開始終了位置のインデクス。</returns>
		public (int start, int end) AppendLine(string str = "")
		{
			return AppendControlCodes(str, "\r\n", 2);
		}

		/// <summary>
		/// 文字列を詰め込み、末尾で改パラグラフを行う。
		/// Word.Document.Content.Textへ変換した時に[SP][CR]で改パラグラフが表現されるように調整する。
		/// </summary>
		/// <param name="str">空文字の場合は空白1文字が追加される。</param>
		/// <returns>与えられた文字列の開始終了位置のインデクス。</returns>
		public (int start, int end) AppendParagraph(string str = "")
		{
			return AppendControlCodes(str, "\r\n\r\n", 2);
		}

		/// <summary>
		/// 文字列と一緒に改行コードを詰め込み、指定されたサイズだけIndexを増やす関数。
		/// </summary>
		/// <param name="str"></param>
		/// <param name="controlCode"></param>
		/// <param name="controlCodeSize"></param>
		/// <returns></returns>
		public (int start, int end) AppendControlCodes(string str, string controlCode, int controlCodeSize)
		{
			int start = 0;
			int end = 0;

			// 改行コード除去
			string appendee = DeleteCrAndLf(str);
			_lineBuilder.Append(appendee);
			string line = _lineBuilder.ToString();

			start = _currentIndex;

			switch (GetLineType(line))
			{
				case LineType.Empty:
					// NOTE: 詰め込むべきものが何も無い場合、スペース1つを挿入して改行する。
					_sentenceBuilder.Append(" ");
					_currentIndex += 1;
					break;
				case LineType.OnlySpaces:
					// NOTE: スペースだけで改行する場合、1文字か2文字以上かで挙動が変わる。
					if (CountEndSpace(line) == 1)
					{
						_sentenceBuilder.Append(line);
						_currentIndex += appendee.Length;
					}
					else
					{
						_sentenceBuilder.Append(line);
						// NOTE: スペースだけの2文字以上スペースは最後の1文字分改行に食われる。
						_currentIndex += appendee.Length - 1;
					}
					break;
				case LineType.EndWithSpaces:
					_sentenceBuilder.Append(line);
					// NOTE: 通常文字列の末尾にスペースが含まれる場合、最後の1文字分改行に食われる。
					_currentIndex += appendee.Length - 1;
					break;
				case LineType.EndWithNoSpace:
					_sentenceBuilder.Append(line);
					_currentIndex += appendee.Length;
					break;
				default:
					return (-1, -1);
			}

			end = _currentIndex;

			// NOTE: 与える制御文字 => Word.Content.Textでの表現
			// [CR][LF] => [SP][VT]
			// [CR][LF] => [SP][CR]
			_sentenceBuilder.Append(controlCode);
			_currentIndex += controlCodeSize;

			// 改行ごとに行バッファはクリアする必要がある。
			_lineBuilder.Clear();

			// Indexは0始まりなのでendは1Index分引く必要がある。
			return (start, end - 1);
		}

		/// <summary>
		/// 蓄積した文字列を1つの文字列にして返す。
		/// </summary>
		/// <returns></returns>
		public string ToString()
		{
			return _sentenceBuilder.ToString();
		}
	}
}
