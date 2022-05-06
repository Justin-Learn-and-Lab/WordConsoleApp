using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordLetterLabConsoleApp
{
	internal class Program
	{
		private static void Main(string[] args)
		{
			var path = @$"{DateTime.Now:yyyyMMddHHmmss}.docx";

			string[] address =
			{
				"新竹科學園區力行六路8號",
				"新竹科學園區園區二路168號",
				"台南科學園區南科北路1號之1",
				"中部科學園區科雅六路1號",
				"台南科學園區北園二路8號",
				"新竹科學園區研新一路9號",
			};

			try
			{
				using (var wordDoc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
				{
					var mainDoc = wordDoc.AddMainDocumentPart();
					mainDoc.Document = new Document();

					var body = new Body();
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

		private static IEnumerable<Paragraph> CreateParagraphs(string[] contents)
		{
			Func<ParagraphProperties> createParagraphPropertiesFunc = () =>
				new ParagraphProperties(new Justification {Val = JustificationValues.Center});
			Func<RunProperties> createRunPropertiesFunc = () =>
				new RunProperties
				{
					FontSize = new FontSize {Val = "36"},
					RunFonts = new RunFonts {EastAsia = "標楷體", Ascii = "Times New Roman"},
				};
			Func<Paragraph> createBreakPageFunc = () => new Paragraph
			(
				new Run(new ParagraphProperties(new WidowControl()),
					new Break {Type = BreakValues.Page})
			);

			for (var i = 0; i < contents.Length - 1; i++)
			{
				yield return new Paragraph
				(
					createParagraphPropertiesFunc(),
					new Run
					(
						createRunPropertiesFunc(),
						new Text(contents[i])
					)
				);
				yield return new Paragraph
				(
					createParagraphPropertiesFunc(),
					new Run
					(
						createRunPropertiesFunc(),
						new Text("台灣積體電路製造股份有限公司 張忠謀 收")
					)
				);
				yield return createBreakPageFunc();
			}

			yield return new Paragraph
			(
				createParagraphPropertiesFunc(),
				new Run
				(
					createRunPropertiesFunc(),
					new Text(contents[contents.Length - 1])
				)
			);
			yield return new Paragraph
			(
				createParagraphPropertiesFunc(),
				new Run
				(
					createRunPropertiesFunc(),
					new Text("台灣積體電路製造股份有限公司 張忠謀 收")
				)
			);
		}

		private static SectionProperties CreateSectionProperties()
		{
			var pageSize = new PageSize
			{
				Width = 12814,
				Height = 6804,
				Code = 6,
				Orient = PageOrientationValues.Landscape,
			};

			var pageMargin = new PageMargin
			{
				Left = 284,
				Top = 284,
				Right = 284,
				Bottom = 284,
				Header = 851,
				Footer = 992,
				Gutter = 0,
			};

			var columns = new Columns {Space = "425"};

			var vAlign = new VerticalTextAlignmentOnPage {Val = VerticalJustificationValues.Center};

			var docGrid = new DocGrid {Type = DocGridValues.LinesAndChars, LinePitch = 360};

			return new SectionProperties(pageSize, pageMargin, columns, vAlign, docGrid);
		}
	}
}