using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace WordConsoleApp
{
	class Program
	{
		static void Main(string[] args)
		{
			string path = @$"{DateTime.Now:yyyyMMddHHmmss}.docx";

			string[] address = new string[]
			{
				"高雄市左營區華夏路390號",
				"高雄市苓雅區成功一路134號",
				"高雄市旗津區中洲三路817號",
				"高雄市鹽埕區五福四路133號",
				"高雄市楠梓區楠梓新路247號",
				"高雄市左營區實踐路43號",
			};

			try
			{
				using (var wordDoc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
				{
					var mainDoc = wordDoc.AddMainDocumentPart();
					mainDoc.Document = new Document();

					Body body = new Body();
					body.Append(CreateParagraphs(address));
					body.Append(CreateSectionProperties());

					mainDoc.Document.Append(body);
					mainDoc.Document.Save();
					wordDoc.Clone();
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
			}
		}

		static IEnumerable<Paragraph> CreateParagraphs(string[] contents)
		{
			for (int i = 0; i < contents.Length - 1; i++)
			{
				yield return new Paragraph
				(
					new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
					new Run(new Text(contents[i]))
				);
				yield return new Paragraph
				(
					new Run(new ParagraphProperties(new WidowControl()), new Break() { Type = BreakValues.Page })
				);
			}
			yield return new Paragraph
			(
				new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
				new Run(new Text(contents[contents.Length - 1]))
			);
		}

		static SectionProperties CreateSectionProperties()
		{
			PageSize pageSize = new PageSize
			{
				Width = 12814,
				Height = 6804,
				Code = 6,
				Orient = PageOrientationValues.Landscape
			};

			var pageMargin = new PageMargin()
			{
				Left = 284,
				Top = 284,
				Right = 284,
				Bottom = 284,
				Header = 851,
				Footer = 992,
				Gutter = 0
			};

			var columns = new Columns { Space = "425" };

			var vAlign = new Word.VerticalTextAlignmentOnPage { Val = VerticalJustificationValues.Center };

			var docGrid = new Word.DocGrid { Type = DocGridValues.LinesAndChars, LinePitch = 360 };

			return new SectionProperties(pageSize, pageMargin, columns, vAlign, docGrid);
		}
	}
}
