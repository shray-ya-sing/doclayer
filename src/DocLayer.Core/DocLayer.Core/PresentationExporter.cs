using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;
using OpenXMLExtensions;
using System.Reflection.Metadata;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace DocLayer.Core
{
    /// <summary>
    /// Converts a slide JSON tree (OOXML-schema format) into a .pptx file.
    ///
    /// Extension method usage map:
    ///   SolidFillExtensions      – SetHexFill on every A.SolidFill
    ///   ShapePropertiesExtensions – SetHorizontalPosition / SetVerticalPosition /
    ///                               SetWidth / SetHeight / SetPresetGeometry /
    ///                               SetHexFill / SetOutlineWidth / SetOutlineHexFill
    ///   GraphicFrameExtensions   – SetHorizontalPosition / SetVerticalPosition /
    ///                               SetWidth / SetHeight on chart + table frames
    ///   RunExtensions            – SetRunEnglish / SetRunSize / SetRunBold /
    ///                               SetRunItalic / SetRunHexFill / AddText
    ///   ParagraphExtensions      – SetAlignLeft / SetAlignCenter / SetAlignRight /
    ///                               SetEndProps / SetText
    ///   TextBodyExtensions       – SetBasicBodyProperties / SetNoAutofit
    ///   TableCellExtensions      – SetHexFill
    ///   ShapeTreeExtensions      – GetShapeId / GetShapeNumber
    ///   SlidePartExtensions      – AddImagePartFromStream (pic elements)
    ///   ShapeStyleExtensions     – SetDefaultReferences
    /// </summary>
    public class PresentationExporter
    {
        // ─────────────────────────────────────────────────────────────────────
        // Public API
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// Deserialises <paramref name="json"/> and writes a .pptx to
        /// <paramref name="outputPath"/>.
        /// </summary>
        public void ExportFromJson(string json, string outputPath)
        {
            var schema = Deserialise(json);
            // PresentationHelperMethods.CreatePresentation creates the full ZIP
            // scaffolding: master, layout, theme, and a blank first slide.
            using var doc = PresentationHelperMethods.CreatePresentation(outputPath);
            PopulateDocument(doc, schema);
            doc.PresentationPart!.Presentation.Save();
        }

        /// <summary>
        /// Exports an array of json slide schema to pptx file
        /// </summary>
        /// <param name="slideJsonArray"></param>
        /// <returns></returns>
        public byte[] ExportPptxBytesFromJsonArray(List<SlideSchema> slideSchemaArray)
        {
            // Guard clause: Don't even create a file if there's no data
            if (slideSchemaArray == null || slideSchemaArray.Count() == 0)
                return [];

            var tmp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".pptx");

            try
            {

                using (var doc = PresentationHelperMethods.CreatePresentation(tmp))
                {
                    // 2. Populate the document using your multi-slide logic
                    PopulateDocumentMultipleSlides(doc, slideSchemaArray);

                    // Note: Ensure PopulateDocumentMultipleSlides or this block 
                    // closes/disposes the doc before ReadAllBytes is called.
                    // 'using' handles this, but the Save() inside Populate is vital.
                }

                return File.ReadAllBytes(tmp);
            }
            finally
            {
                if (File.Exists(tmp)) File.Delete(tmp);
            }
        }


        /// <summary>
        /// Reads JSON from <paramref name="jsonPath"/> and writes a .pptx to
        /// <paramref name="outputPath"/>.
        /// </summary>
        public void ExportFromFile(string jsonPath, string outputPath)
        {
            if (!File.Exists(jsonPath))
                throw new FileNotFoundException("Slide JSON not found.", jsonPath);

            ExportFromJson(File.ReadAllText(jsonPath), outputPath);
        }

        /// <summary>
        /// Returns the presentation as an in-memory byte array.
        /// Useful for streaming directly from an HTTP response without touching
        /// the filesystem.
        /// </summary>
        public byte[] ExportToBytes(string json)
        {
            // Write to a temp file then read back, so we can reuse the same
            // CreatePresentation scaffolding path without duplicating it.
            var tmp = Path.GetTempFileName() + ".pptx";
            try
            {
                ExportFromJson(json, tmp);
                return File.ReadAllBytes(tmp);
            }
            finally
            {
                if (File.Exists(tmp)) File.Delete(tmp);
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        // Deserialisation
        // ─────────────────────────────────────────────────────────────────────

        private static SlideSchema Deserialise(string json) =>
            JsonSerializer.Deserialize<SlideSchema>(json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new ArgumentException("Invalid or empty slide JSON.");

        // ─────────────────────────────────────────────────────────────────────
        // Document-level population
        // ─────────────────────────────────────────────────────────────────────

        private static void PopulateDocument(PresentationDocument doc, SlideSchema schema)
        {
            var slideDef = schema.Slide;
            var pPart = doc.PresentationPart!;
            var slidePart = pPart.SlideParts.First();  // created by CreatePresentation
            var oSlide = slidePart.Slide;

            // ── Slide dimensions ──────────────────────────────────────────
            // PresentationExtensions.SetSlideSizeWidescreen sets 12192000×6858000;
            // here we honour whatever the JSON specifies (already in EMU).
            var slideSize = pPart.Presentation.SlideSize;
            if (slideSize != null)
            {
                slideSize.Cx = (Int32Value)(int)slideDef.Width;
                slideSize.Cy = (Int32Value)(int)slideDef.Height;
            }

            // ── Background ────────────────────────────────────────────────
            if (slideDef.Background?.Fill != null)
                ApplyBackground(oSlide, slideDef.Background.Fill);

            // ── Shape tree — remove default title placeholder ─────────────
            var tree = oSlide.CommonSlideData!.ShapeTree!;
            foreach (var ph in tree.Elements<P.Shape>().ToList())
                ph.Remove();

            // ── Elements ──────────────────────────────────────────────────
            foreach (var el in slideDef.Elements ?? [])
            {
                switch (el.Type?.ToLowerInvariant())
                {
                    case "sp":
                        tree.AppendChild(BuildShape(tree, el));
                        break;
                    case "cxnsp":
                        tree.AppendChild(BuildConnector(tree, el));
                        break;
                    case "chart":
                        tree.AppendChild(BuildChartFrame(slidePart, tree, el));
                        break;
                    case "table":
                        tree.AppendChild(BuildTableFrame(tree, el));
                        break;
                    case "pic":
                        BuildPic(slidePart, tree, el);
                        break;
                }
            }

            oSlide.Save();
        }

        private static void PopulateDocumentMultipleSlides(PresentationDocument doc, List<SlideSchema> schemas)
        {
            var pPart = doc.PresentationPart!;
            var firstSlideDef = schemas[0].Slide;

            // 1. Set global slide dimensions (one-time operation)
            var slideSize = pPart.Presentation.SlideSize;
            if (slideSize != null)
            {
                slideSize.Cx = (Int32Value)(int)firstSlideDef.Width;
                slideSize.Cy = (Int32Value)(int)firstSlideDef.Height;
            }

            // 2. Iterate through schemas and assign/create slides
            for (int i = 0; i < schemas.Count(); i++)
            {
                SlidePart slidePart;

                if (i == 0)
                {
                    // Use the slide created by default in your helper method
                    slidePart = pPart.SlideParts.First();
                }
                else
                {
                    // Use your new extension method to append a fresh slide
                    slidePart = doc.AddNewSlide();
                }

                // 3. Populate the specific slide
                PopulateSlide(slidePart, schemas[i]);
            }

            doc.PresentationPart?.Presentation.Save();
        }

        private static void PopulateSlide(SlidePart slidePart, SlideSchema schema)
        {
            var slideDef = schema.Slide;
            var oSlide = slidePart.Slide;

            // ── Background ────────────────────────────────────────────────
            if (slideDef.Background?.Fill != null)
                ApplyBackground(oSlide, slideDef.Background.Fill);

            // ── Shape tree — remove default title placeholder ─────────────
            var tree = oSlide.CommonSlideData!.ShapeTree!;
            foreach (var ph in tree.Elements<P.Shape>().ToList())
                ph.Remove();

            // ── Elements ──────────────────────────────────────────────────
            foreach (var el in slideDef.Elements ?? [])
            {
                switch (el.Type?.ToLowerInvariant())
                {
                    case "sp":
                        tree.AppendChild(BuildShape(tree, el));
                        break;
                    case "cxnsp":
                        tree.AppendChild(BuildConnector(tree, el));
                        break;
                    case "chart":
                        tree.AppendChild(BuildChartFrame(slidePart, tree, el));
                        break;
                    case "table":
                        tree.AppendChild(BuildTableFrame(tree, el));
                        break;
                    case "pic":
                        BuildPic(slidePart, tree, el);
                        break;
                }
            }

            oSlide.Save();
        }

        // ─────────────────────────────────────────────────────────────────────
        // Background
        // ─────────────────────────────────────────────────────────────────────

        private static void ApplyBackground(Slide slide, FillDefinition fill)
        {
            var bg = new P.Background();
            var bgPr = new P.BackgroundProperties();

            if (fill.Type?.ToLowerInvariant() == "solid" && !string.IsNullOrEmpty(fill.Color))
            {
                var sf = new A.SolidFill();
                sf.SetHexFill(Hex(fill.Color));   // SolidFillExtensions
                bgPr.AppendChild(sf);
            }
            else
            {
                bgPr.AppendChild(new A.NoFill());
            }

            bg.AppendChild(bgPr);
            // OOXML schema requires <p:bg> before <p:cSld>
            slide.InsertBefore(bg, slide.CommonSlideData);
        }

        // ─────────────────────────────────────────────────────────────────────
        // Shape  <p:sp>
        // ─────────────────────────────────────────────────────────────────────

        private static P.Shape BuildShape(P.ShapeTree tree, ElementDefinition el)
        {
            var sp = new P.Shape();

            // ShapeTreeExtensions.GetShapeId / GetShapeNumber keep IDs
            // consistent with every other element already in the tree.
            UInt32Value id = el.Id > 0
                ? (UInt32Value)(uint)el.Id
                : tree.GetShapeId();                    // ShapeTreeExtensions
            string num = tree.GetShapeNumber();         // ShapeTreeExtensions

            sp.AppendChild(new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = el.Name ?? $"sp{num}" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()));

            // ── Shape properties ────────────────────────────────────────
            var spPr = new P.ShapeProperties();
            sp.AppendChild(spPr);

            // Geometry — ShapePropertiesExtensions
            spPr.SetPresetGeometry(A.ShapeTypeValues.Rectangle);

            // Position & size — ShapePropertiesExtensions (EMU overloads)
            ApplyPositionEmu(spPr, el.Position);

            // Fill — ShapePropertiesExtensions
            if (el.Fill?.Type?.ToLowerInvariant() == "solid" && !string.IsNullOrEmpty(el.Fill.Color))
                spPr.SetHexFill(Hex(el.Fill.Color));   // ShapePropertiesExtensions
            else
                spPr.AppendChild(new A.NoFill());

            // Outline — ShapePropertiesExtensions
            ApplyOutline(spPr, el.Border, el.Line);

            // ── Shape style (needed for theme-colour references) ─────────
            var style = new P.ShapeStyle();
            style.SetDefaultReferences();               // ShapeStyleExtensions
            sp.AppendChild(style);

            // ── Text body ───────────────────────────────────────────────
            sp.AppendChild(el.Text?.Body != null
                ? BuildTextBody(el.Text.Body)
                : EmptyTextBody());

            return sp;
        }

        // ─────────────────────────────────────────────────────────────────────
        // Connector  <p:cxnSp>
        // ─────────────────────────────────────────────────────────────────────

        private static P.ConnectionShape BuildConnector(P.ShapeTree tree, ElementDefinition el)
        {
            var cxn = new P.ConnectionShape();

            UInt32Value id = el.Id > 0
                ? (UInt32Value)(uint)el.Id
                : tree.GetShapeId();                    // ShapeTreeExtensions
            string num = tree.GetShapeNumber();

            cxn.AppendChild(new P.NonVisualConnectionShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = el.Name ?? $"cxn{num}" },
                new P.NonVisualConnectorShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()));

            var spPr = new P.ShapeProperties();
            cxn.AppendChild(spPr);

            // ShapePropertiesExtensions
            spPr.SetPresetGeometry(A.ShapeTypeValues.Line);
            ApplyPositionEmu(spPr, el.Position);

            // For connectors the outline IS the visible stroke; build it
            // manually so we can attach optional arrowheads.
            if (el.Line != null)
            {
                var ln = new A.Outline
                { Width = el.Line.Width > 0 ? (Int32Value)el.Line.Width : 9525 };

                var sf = new A.SolidFill();
                sf.SetHexFill(Hex(el.Line.Color));      // SolidFillExtensions
                ln.AppendChild(sf);

                if (el.HeadEnd != null)
                    ln.AppendChild(new A.HeadEnd { Type = ArrowType(el.HeadEnd.Type) });
                if (el.TailEnd != null)
                    ln.AppendChild(new A.TailEnd { Type = ArrowType(el.TailEnd.Type) });

                spPr.AppendChild(ln);
            }

            // Shape style
            var style = new P.ShapeStyle();
            style.SetDefaultReferences();               // ShapeStyleExtensions
            cxn.AppendChild(style);

            return cxn;
        }

        // ─────────────────────────────────────────────────────────────────────
        // Picture  <p:pic>
        // Uses SlidePartExtensions.AddImagePartFromStream + ShapeTreeExtensions.AddPicture
        // ─────────────────────────────────────────────────────────────────────

        private static void BuildPic(
            SlidePart slidePart, P.ShapeTree tree, ElementDefinition el)
        {
            if (el.ImageData == null) return;

            // SlidePartExtensions — registers the image bytes as an ImagePart
            // and returns the relationship id.
            using var ms = new MemoryStream(el.ImageData);
            string relId = slidePart.AddImagePartFromStream(ms); // SlidePartExtensions

            var pos = el.Position;
            if (pos != null)
            {
                // ShapeTreeExtensions.AddPicture (decimal EMU overload)
                // The overload that accepts height + width is the one we want.
                tree.AddPicture(
                    relId,
                    height: (decimal)pos.Cy,
                    width: (decimal)pos.Cx,
                    hpos: (decimal)pos.X,
                    vpos: (decimal)pos.Y);
            }
            else
            {
                tree.AddPicture(relId, hpos: 0m, vpos: 0m);
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        // Chart graphic frame  <p:graphicFrame>
        // ─────────────────────────────────────────────────────────────────────

        private static P.GraphicFrame BuildChartFrame(
            SlidePart slidePart, P.ShapeTree tree, ElementDefinition el)
        {
            var chartPart = slidePart.AddNewPart<ChartPart>();
            PopulateChartPart(chartPart, el);
            string relId = slidePart.GetIdOfPart(chartPart);

            UInt32Value id = el.Id > 0
                ? (UInt32Value)(uint)el.Id
                : tree.GetShapeId();                    // ShapeTreeExtensions
            string num = tree.GetShapeNumber();

            var frame = new P.GraphicFrame();

            frame.AppendChild(new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = el.Name ?? $"chart{num}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()));

            // Seed a Transform so GraphicFrameExtensions can find and mutate it.
            frame.AppendChild(new P.Transform(
                new A.Offset { X = 0, Y = 0 },
                new A.Extents { Cx = 0, Cy = 0 }));

            // GraphicFrameExtensions — same defensive pattern as the rest of the library
            if (el.Position != null)
            {
                frame.SetHorizontalPosition((Int64Value)el.Position.X);  // GraphicFrameExtensions
                frame.SetVerticalPosition((Int64Value)el.Position.Y);    // GraphicFrameExtensions
                frame.SetWidth((Int64Value)el.Position.Cx);              // GraphicFrameExtensions
                frame.SetHeight((Int64Value)el.Position.Cy);             // GraphicFrameExtensions
            }

            var gData = new A.GraphicData
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };
            gData.AppendChild(new C.ChartReference { Id = relId });
            frame.AppendChild(new A.Graphic(gData));

            return frame;
        }

        // ─────────────────────────────────────────────────────────────────────
        // Chart content (ChartPart XML)
        // ─────────────────────────────────────────────────────────────────────

        private static void PopulateChartPart(ChartPart chartPart, ElementDefinition el)
        {
            var cs = new C.ChartSpace();
            cs.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            cs.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            cs.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            cs.AppendChild(new C.Date1904 { Val = false });
            cs.AppendChild(new C.RoundedCorners { Val = false });

            var chart = new C.Chart();
            chart.AppendChild(new C.AutoTitleDeleted { Val = true });

            var plotArea = new C.PlotArea();

            // Plot area background — SolidFillExtensions
            if (el.PlotArea?.Fill?.Type?.ToLowerInvariant() == "solid"
                && !string.IsNullOrEmpty(el.PlotArea.Fill.Color))
            {
                var paSp = new C.ShapeProperties();
                var paSf = new A.SolidFill();
                paSf.SetHexFill(Hex(el.PlotArea.Fill.Color));  // SolidFillExtensions
                paSp.AppendChild(paSf);
                paSp.AppendChild(new A.Outline(new A.NoFill()));
                plotArea.AppendChild(paSp);
            }

            var chartType = el.ChartType?.ToLowerInvariant() ?? "linechart";
            if (chartType == "linechart")
                plotArea.AppendChild(BuildLineChart(el));
            else if (chartType == "barchart")
                plotArea.AppendChild(BuildBarChart(el));

            if (el.Axes != null)
            {
                plotArea.AppendChild(BuildCategoryAxis(100u, 200u, el.Axes.CatAx));
                plotArea.AppendChild(BuildValueAxis(200u, 100u, el.Axes.ValAx));
            }

            chart.AppendChild(plotArea);

            if (el.Legend?.Visible == true)
            {
                var leg = new C.Legend();
                leg.AppendChild(new C.LegendPosition { Val = C.LegendPositionValues.Bottom });
                chart.AppendChild(leg);
            }

            chart.AppendChild(new C.PlotVisibleOnly { Val = true });
            cs.AppendChild(chart);

            // Chart-space background (mirrors plot area so dark panels look correct)
            if (el.PlotArea?.Fill?.Type?.ToLowerInvariant() == "solid"
                && !string.IsNullOrEmpty(el.PlotArea.Fill.Color))
            {
                var csSp = new C.ShapeProperties();
                var csSf = new A.SolidFill();
                csSf.SetHexFill(Hex(el.PlotArea.Fill.Color));  // SolidFillExtensions
                csSp.AppendChild(csSf);
                csSp.AppendChild(new A.Outline(new A.NoFill()));
                cs.AppendChild(csSp);
            }

            chartPart.ChartSpace = cs;
            chartPart.ChartSpace.Save();
        }

        // ── Line chart ──────────────────────────────────────────────────────

        private static C.LineChart BuildLineChart(ElementDefinition el)
        {
            var lc = new C.LineChart();
            lc.AppendChild(new C.Grouping { Val = C.GroupingValues.Standard });
            lc.AppendChild(new C.VaryColors { Val = false });

            uint idx = 0;
            foreach (var s in el.Series ?? [])
            {
                var ser = new C.LineChartSeries();
                ser.AppendChild(new C.Index { Val = idx });
                ser.AppendChild(new C.Order { Val = idx });
                ser.AppendChild(new C.SeriesText(StringRef([s.Name ?? ""])));

                // Marker
                var marker = new C.Marker();
                marker.AppendChild(new C.Symbol { Val = C.MarkerStyleValues.Circle });
                marker.AppendChild(new C.Size
                { Val = (ByteValue)(byte)(s.MarkerSize > 0 ? s.MarkerSize : 5) });
                if (!string.IsNullOrEmpty(s.MarkerColor))
                {
                    var mSp = new C.ShapeProperties();
                    var mSf = new A.SolidFill();
                    mSf.SetHexFill(Hex(s.MarkerColor));        // SolidFillExtensions
                    mSp.AppendChild(mSf);
                    var mLn = new A.Outline();
                    var mLnSf = new A.SolidFill();
                    mLnSf.SetHexFill(Hex(s.MarkerColor));      // SolidFillExtensions
                    mLn.AppendChild(mLnSf);
                    mSp.AppendChild(mLn);
                    marker.AppendChild(mSp);
                }
                ser.AppendChild(marker);

                // Line colour
                if (!string.IsNullOrEmpty(s.Color))
                {
                    var sSp = new C.ShapeProperties();
                    var ln = new A.Outline { Width = 19050 };
                    var sf = new A.SolidFill();
                    sf.SetHexFill(Hex(s.Color));               // SolidFillExtensions
                    ln.AppendChild(sf);
                    sSp.AppendChild(ln);
                    ser.AppendChild(sSp);
                }

                ser.AppendChild(new C.Smooth { Val = s.Smooth });
                AppendSeriesData(ser, s);
                lc.AppendChild(ser);
                idx++;
            }

            lc.AppendChild(new C.AxisId { Val = 100 });
            lc.AppendChild(new C.AxisId { Val = 200 });
            return lc;
        }

        // ── Bar chart ───────────────────────────────────────────────────────

        private static C.BarChart BuildBarChart(ElementDefinition el)
        {
            var bc = new C.BarChart();
            bc.AppendChild(new C.BarDirection
            {
                Val = el.BarDir?.ToLowerInvariant() == "bar"
                    ? C.BarDirectionValues.Bar
                    : C.BarDirectionValues.Column
            });
            bc.AppendChild(new C.Grouping { Val = C.GroupingValues.Standard });
            bc.AppendChild(new C.VaryColors { Val = false });

            uint idx = 0;
            foreach (var s in el.Series ?? [])
            {
                var ser = new C.BarChartSeries();
                ser.AppendChild(new C.Index { Val = idx });
                ser.AppendChild(new C.Order { Val = idx });
                ser.AppendChild(new C.SeriesText(StringRef([s.Name ?? ""])));

                if (!string.IsNullOrEmpty(s.Color))
                {
                    var sSp = new C.ShapeProperties();
                    var sf = new A.SolidFill();
                    sf.SetHexFill(Hex(s.Color));               // SolidFillExtensions
                    sSp.AppendChild(sf);
                    ser.AppendChild(sSp);
                }

                // negativeColor in the JSON schema maps to InvertIfNegative in OOXML
                if (!string.IsNullOrEmpty(s.NegativeColor))
                    ser.AppendChild(new C.InvertIfNegative { Val = true });

                if (el.DataLabels?.Visible == true)
                    ser.AppendChild(BuildDataLabels(el.DataLabels));

                AppendSeriesData(ser, s);
                bc.AppendChild(ser);
                idx++;
            }

            bc.AppendChild(new C.AxisId { Val = 100 });
            bc.AppendChild(new C.AxisId { Val = 200 });
            return bc;
        }

        // ── Data labels ─────────────────────────────────────────────────────

        private static C.DataLabels BuildDataLabels(DataLabelsDefinition dl)
        {
            var dLbls = new C.DataLabels();
            dLbls.AppendChild(new C.ShowLegendKey { Val = false });
            dLbls.AppendChild(new C.ShowValue { Val = true });
            dLbls.AppendChild(new C.ShowCategoryName { Val = false });
            dLbls.AppendChild(new C.ShowSeriesName { Val = false });
            dLbls.AppendChild(new C.ShowPercent { Val = false });
            dLbls.AppendChild(new C.ShowBubbleSize { Val = false });

            if (!string.IsNullOrEmpty(dl.Color))
            {
                // RunExtensions for label text colour
                var run = new A.Run(
                    new A.RunProperties { Language = "en-US" },
                    new A.Text(""));
                run.SetRunSize(dl.FontSize > 0 ? dl.FontSize / 100 : 7); // RunExtensions — pts
                run.SetRunHexFill(Hex(dl.Color));                        // RunExtensions

                dLbls.AppendChild(new C.TextProperties(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(run)));
            }

            return dLbls;
        }

        // ── Axis label text properties ───────────────────────────────────────

        private static C.TextProperties AxisTextProperties(string color, int fontSizeHalfPt)
        {
            // RunExtensions — SetRunSize takes pts; fontSizeHalfPt ÷ 100 = pts
            var run = new A.Run(
                new A.RunProperties { Language = "en-US" },
                new A.Text(""));
            run.SetRunSize(fontSizeHalfPt > 0 ? fontSizeHalfPt / 100 : 8); // RunExtensions
            run.SetRunHexFill(Hex(color));                                  // RunExtensions

            return new C.TextProperties(
                new A.BodyProperties { Rotation = 0 },
                new A.ListStyle(),
                new A.Paragraph(run));
        }

        // ── Series data helpers ──────────────────────────────────────────────

        private static void AppendSeriesData(OpenXmlCompositeElement ser, SeriesDefinition s)
        {
            var pts = s.Points ?? [];
            var labels = pts.Select(p => p.Label ?? "").ToArray();
            var vals = pts.Select(p => p.Value).ToArray();
            ser.AppendChild(new C.CategoryAxisData(StringRef(labels)));
            ser.AppendChild(new C.Values(NumberRef(vals)));
        }

        // ─────────────────────────────────────────────────────────────────────
        // Axes
        // ─────────────────────────────────────────────────────────────────────

        private static C.CategoryAxis BuildCategoryAxis(
            uint axId, uint crossId, AxisDefinition? ax)
        {
            var ca = new C.CategoryAxis();
            ca.AppendChild(new C.AxisId { Val = axId });
            ca.AppendChild(new C.Scaling(
                new C.Orientation { Val = C.OrientationValues.MinMax }));
            ca.AppendChild(new C.Delete { Val = ax?.Visible != true });
            ca.AppendChild(new C.AxisPosition { Val = C.AxisPositionValues.Bottom });
            ca.AppendChild(new C.TickLabelPosition
            { Val = C.TickLabelPositionValues.NextTo });
            ca.AppendChild(new C.CrossingAxis { Val = crossId });
            ca.AppendChild(new C.Crosses { Val = C.CrossesValues.AutoZero });
            ca.AppendChild(new C.AutoLabeled { Val = true });
            ca.AppendChild(new C.LabelAlignment { Val = C.LabelAlignmentValues.Center });
            ca.AppendChild(new C.LabelOffset { Val = 100 });

            if (ax?.Visible == true && !string.IsNullOrEmpty(ax.LabelColor))
                ca.AppendChild(AxisTextProperties(ax.LabelColor, ax.LabelFontSize));

            return ca;
        }

        private static C.ValueAxis BuildValueAxis(
            uint axId, uint crossId, AxisDefinition? ax)
        {
            var va = new C.ValueAxis();
            va.AppendChild(new C.AxisId { Val = axId });

            var scaling = new C.Scaling(
                new C.Orientation { Val = C.OrientationValues.MinMax });
            if (ax?.Min.HasValue == true)
                scaling.AppendChild(new C.MinAxisValue { Val = ax.Min!.Value });
            if (ax?.Max.HasValue == true)
                scaling.AppendChild(new C.MaxAxisValue { Val = ax.Max!.Value });
            va.AppendChild(scaling);

            va.AppendChild(new C.Delete { Val = ax?.Visible != true });
            va.AppendChild(new C.AxisPosition { Val = C.AxisPositionValues.Left });

            if (ax?.GridLine?.Type?.ToLowerInvariant() == "none")
                va.AppendChild(new C.MajorGridlines(
                    new C.ShapeProperties(new A.NoFill(), new A.Outline(new A.NoFill()))));

            if (!string.IsNullOrEmpty(ax?.NumFmt))
                va.AppendChild(new C.NumberingFormat
                { FormatCode = ax!.NumFmt, SourceLinked = false });

            va.AppendChild(new C.TickLabelPosition
            { Val = C.TickLabelPositionValues.NextTo });
            va.AppendChild(new C.CrossingAxis { Val = crossId });
            va.AppendChild(new C.Crosses { Val = C.CrossesValues.AutoZero });
            va.AppendChild(new C.CrossBetween { Val = C.CrossBetweenValues.Between });

            if (ax?.Visible == true && !string.IsNullOrEmpty(ax.LabelColor))
                va.AppendChild(AxisTextProperties(ax.LabelColor, ax.LabelFontSize));

            return va;
        }

        // ─────────────────────────────────────────────────────────────────────
        // Table graphic frame  <p:graphicFrame> + <a:tbl>
        // ─────────────────────────────────────────────────────────────────────

        private static P.GraphicFrame BuildTableFrame(P.ShapeTree tree, ElementDefinition el)
        {
            UInt32Value id = el.Id > 0
                ? (UInt32Value)(uint)el.Id
                : tree.GetShapeId();                    // ShapeTreeExtensions
            string num = tree.GetShapeNumber();

            var frame = new P.GraphicFrame();

            frame.AppendChild(new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = el.Name ?? $"tbl{num}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()));

            // Seed a Transform so GraphicFrameExtensions can find and mutate it.
            frame.AppendChild(new P.Transform(
                new A.Offset { X = 0, Y = 0 },
                new A.Extents { Cx = 0, Cy = 0 }));

            // GraphicFrameExtensions
            if (el.Position != null)
            {
                frame.SetHorizontalPosition((Int64Value)el.Position.X);  // GraphicFrameExtensions
                frame.SetVerticalPosition((Int64Value)el.Position.Y);    // GraphicFrameExtensions
                frame.SetWidth((Int64Value)el.Position.Cx);              // GraphicFrameExtensions
                frame.SetHeight((Int64Value)el.Position.Cy);             // GraphicFrameExtensions
            }

            var gData = new A.GraphicData
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };
            gData.AppendChild(BuildTable(el));
            frame.AppendChild(new A.Graphic(gData));

            return frame;
        }

        private static A.Table BuildTable(ElementDefinition el)
        {
            var tbl = new A.Table();
            tbl.AppendChild(new A.TableProperties { FirstRow = false, BandRow = false });

            var grid = new A.TableGrid();
            foreach (var col in el.Columns ?? [])
                grid.AppendChild(new A.GridColumn
                { Width = col.Width > 0 ? (Int64Value)col.Width : 1028700L });
            tbl.AppendChild(grid);

            foreach (var row in el.Rows ?? [])
            {
                var tr = new A.TableRow
                { Height = row.Height > 0 ? (Int64Value)row.Height : 200000L };
                foreach (var cell in row.Cells ?? [])
                    tr.AppendChild(BuildTableCell(cell));
                tbl.AppendChild(tr);
            }

            return tbl;
        }

        private static A.TableCell BuildTableCell(CellDefinition cell)
        {
            var tc = new A.TableCell();

            // ── Text body ─────────────────────────────────────────────────
            var para = new A.Paragraph();

            // ParagraphExtensions for alignment
            switch (cell.Alignment?.ToLowerInvariant())
            {
                case "ctr":
                case "center": para.SetAlignCenter(); break;   // ParagraphExtensions
                case "r":
                case "right": para.SetAlignRight(); break;   // ParagraphExtensions
                default: para.SetAlignLeft(); break;   // ParagraphExtensions
            }

            // RunExtensions for all run-level formatting
            var run = new A.Run();
            run.SetRunEnglish();                                // RunExtensions
            // SetRunSize takes points; JSON stores half-points (÷100 = pts)
            run.SetRunSize(cell.FontSize > 0 ? cell.FontSize / 100 : 8); // RunExtensions
            if (cell.Bold) run.SetRunBold();                 // RunExtensions
            if (cell.Italic) run.SetRunItalic();               // RunExtensions
            if (!string.IsNullOrEmpty(cell.Color))
                run.SetRunHexFill(Hex(cell.Color));            // RunExtensions
            run.AddText(cell.Text ?? "");                      // RunExtensions

            para.AppendChild(run);
            para.SetEndProps();                                // ParagraphExtensions

            var txBody = new A.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                para);
            tc.AppendChild(txBody);

            // ── Cell properties + fill ────────────────────────────────────
            // TableCellExtensions.SetHexFill requires the cell to already have
            // a TableCellProperties child, so we add it first.
            tc.AppendChild(new A.TableCellProperties());

            if (cell.Fill?.Type?.ToLowerInvariant() == "solid"
                && !string.IsNullOrEmpty(cell.Fill.Color))
            {
                tc.SetHexFill(Hex(cell.Fill.Color));           // TableCellExtensions
            }

            return tc;
        }

        // ─────────────────────────────────────────────────────────────────────
        // Text body  <p:txBody>
        // ─────────────────────────────────────────────────────────────────────

        private static P.TextBody BuildTextBody(TextBodyDefinition body)
        {
            var txBody = new P.TextBody();

            // BodyProperties
            var bodyPr = new A.BodyProperties { Wrap = A.TextWrappingValues.Square };
            bodyPr.Anchor = body.Anchor?.ToLowerInvariant() switch
            {
                "t" => A.TextAnchoringTypeValues.Top,
                "ctr" => A.TextAnchoringTypeValues.Center,
                "b" => A.TextAnchoringTypeValues.Bottom,
                _ => A.TextAnchoringTypeValues.Top
            };
            txBody.AppendChild(bodyPr);
            txBody.AppendChild(new A.ListStyle());

            // Autofit — BodyPropertiesExtensions (via TextBodyExtensions)
            if (body.Autofit == true)
                txBody.SetShapeAutofit();   // TextBodyExtensions → BodyPropertiesExtensions
            else
                txBody.SetNoAutofit();      // TextBodyExtensions → BodyPropertiesExtensions

            foreach (var paraDef in body.Paragraphs ?? [])
                txBody.AppendChild(BuildParagraph(paraDef));

            return txBody;
        }

        private static A.Paragraph BuildParagraph(ParagraphDefinition paraDef)
        {
            var para = new A.Paragraph();

            // ParagraphExtensions
            switch (paraDef.Alignment?.ToLowerInvariant())
            {
                case "ctr":
                case "center": para.SetAlignCenter(); break;   // ParagraphExtensions
                case "r":
                case "right": para.SetAlignRight(); break;   // ParagraphExtensions
                default: para.SetAlignLeft(); break;   // ParagraphExtensions
            }

            if (paraDef.LineSpacing > 0)
                para.SetLineSpacing(paraDef.LineSpacing);      // ParagraphExtensions

            foreach (var runDef in paraDef.Runs ?? [])
                para.AppendChild(BuildRun(runDef));

            para.SetEndProps();                                // ParagraphExtensions
            return para;
        }

        private static A.Run BuildRun(RunDefinition r)
        {
            var run = new A.Run();

            // RunExtensions for every property
            run.SetRunEnglish();                               // RunExtensions
            run.SetRunSize(r.FontSize > 0 ? r.FontSize / 100 : 10); // RunExtensions — pts
            if (r.Bold) run.SetRunBold();                    // RunExtensions
            if (r.Italic) run.SetRunItalic();                  // RunExtensions

            // FontFace and Baseline have no dedicated extension yet — set directly.
            if (!string.IsNullOrEmpty(r.FontFace) || r.Baseline != 0)
            {
                // RunProperties is created by SetRunEnglish, so it exists now.
                if (!string.IsNullOrEmpty(r.FontFace))
                    run.RunProperties!.AppendChild(
                        new A.LatinFont { Typeface = r.FontFace });
                if (r.Baseline != 0)
                    run.RunProperties!.Baseline = (Int32Value)r.Baseline;
            }

            if (!string.IsNullOrEmpty(r.Color))
                run.SetRunHexFill(Hex(r.Color));               // RunExtensions

            run.AddText(r.Text ?? "");                         // RunExtensions
            return run;
        }

        // ─────────────────────────────────────────────────────────────────────
        // Shared position / outline helpers
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// Applies x, y, cx, cy (all EMU) using ShapePropertiesExtensions,
        /// which handle the Transform2D / Offset / Extents sub-tree defensively.
        /// </summary>
        private static void ApplyPositionEmu(P.ShapeProperties spPr, PositionDefinition? pos)
        {
            if (pos == null) return;
            spPr.SetHorizontalPosition((Int64Value)pos.X);    // ShapePropertiesExtensions
            spPr.SetVerticalPosition((Int64Value)pos.Y);      // ShapePropertiesExtensions
            spPr.SetWidth((Int64Value)pos.Cx);                // ShapePropertiesExtensions
            spPr.SetHeight((Int64Value)pos.Cy);               // ShapePropertiesExtensions
        }

        /// <summary>
        /// Applies border or line stroke to ShapeProperties using
        /// ShapePropertiesExtensions.
        /// SetOutlineWidth takes a double in points (multiply EMU width ÷ 12700).
        /// </summary>
        private static void ApplyOutline(
            P.ShapeProperties spPr,
            BorderDefinition? border,
            LineDefinition? line)
        {
            if (line != null)
            {
                double widthPt = line.Width > 0 ? line.Width / 12700.0 : 0.75;
                spPr.SetOutlineWidth(widthPt);                // ShapePropertiesExtensions
                if (!string.IsNullOrEmpty(line.Color))
                    spPr.SetOutlineHexFill(Hex(line.Color));  // ShapePropertiesExtensions
            }
            else if (border?.Type?.ToLowerInvariant() == "solid"
                     && !string.IsNullOrEmpty(border.Color))
            {
                double widthPt = border.Width > 0 ? border.Width / 12700.0 : 0.75;
                spPr.SetOutlineWidth(widthPt);                // ShapePropertiesExtensions
                spPr.SetOutlineHexFill(Hex(border.Color));    // ShapePropertiesExtensions
            }
            else
            {
                spPr.AppendChild(new A.Outline(new A.NoFill()));
            }
        }

        private static P.TextBody EmptyTextBody() =>
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.EndParagraphRunProperties { Language = "en-US" }));

        // ─────────────────────────────────────────────────────────────────────
        // Chart data reference helpers
        // ─────────────────────────────────────────────────────────────────────

        private static C.StringReference StringRef(string[] items)
        {
            var cache = new C.StringCache();
            cache.AppendChild(new C.PointCount
            { Val = (UInt32Value)(uint)items.Length });
            for (uint i = 0; i < items.Length; i++)
                cache.AppendChild(new C.StringPoint
                { Index = i, NumericValue = new C.NumericValue(items[i]) });

            var sr = new C.StringReference();
            sr.AppendChild(new C.Formula("Sheet1!$A$1"));
            sr.AppendChild(cache);
            return sr;
        }

        private static C.NumberReference NumberRef(double[] values)
        {
            var cache = new C.NumberingCache();
            cache.AppendChild(new C.FormatCode("General"));
            cache.AppendChild(new C.PointCount
            { Val = (UInt32Value)(uint)values.Length });
            for (uint i = 0; i < values.Length; i++)
                cache.AppendChild(new C.NumericPoint
                {
                    Index = i,
                    NumericValue = new C.NumericValue(values[i].ToString("G"))
                });

            var nr = new C.NumberReference();
            nr.AppendChild(new C.Formula("Sheet1!$B$1"));
            nr.AppendChild(cache);
            return nr;
        }

        // ─────────────────────────────────────────────────────────────────────
        // Enum / string converters
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>Normalises a colour to 6-char uppercase hex without #.</summary>
        private static string Hex(string? color) =>
            (color ?? "000000").TrimStart('#').ToUpperInvariant();

        private static A.LineEndValues ArrowType(string? v) =>
            v?.ToLowerInvariant() switch
            {
                "arrow" => A.LineEndValues.Arrow,
                "stealth" => A.LineEndValues.Stealth,
                "diamond" => A.LineEndValues.Diamond,
                "oval" => A.LineEndValues.Oval,
                _ => A.LineEndValues.None
            };
    }
}