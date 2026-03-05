using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace DocLayer.Core
{
    // ─────────────────────────────────────────────────────────────────────────
    // Root
    // ─────────────────────────────────────────────────────────────────────────

    public class SlideSchema
    {
        [JsonPropertyName("slide")]
        public SlideDefinition Slide { get; set; } = new();
    }

    public class SlideDefinition
    {
        [JsonPropertyName("width")]
        public long Width { get; set; } = 9144000;

        [JsonPropertyName("height")]
        public long Height { get; set; } = 5143500;

        [JsonPropertyName("background")]
        public BackgroundDefinition? Background { get; set; }

        [JsonPropertyName("elements")]
        public List<ElementDefinition>? Elements { get; set; }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Elements
    // ─────────────────────────────────────────────────────────────────────────

    public class ElementDefinition
    {
        /// <summary>sp | cxnSp | chart | table | pic</summary>
        [JsonPropertyName("type")]
        public string? Type { get; set; }

        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("position")]
        public PositionDefinition? Position { get; set; }

        [JsonPropertyName("imageData")]
        public byte[]? ImageData { get; set; }

        // ── Shape properties ──────────────────────────────────────────────

        [JsonPropertyName("fill")]
        public FillDefinition? Fill { get; set; }

        [JsonPropertyName("border")]
        public BorderDefinition? Border { get; set; }

        [JsonPropertyName("line")]
        public LineDefinition? Line { get; set; }

        [JsonPropertyName("headEnd")]
        public ArrowEndDefinition? HeadEnd { get; set; }

        [JsonPropertyName("tailEnd")]
        public ArrowEndDefinition? TailEnd { get; set; }

        [JsonPropertyName("text")]
        public TextDefinition? Text { get; set; }

        // ── Chart properties ──────────────────────────────────────────────

        /// <summary>lineChart | barChart | pieChart</summary>
        [JsonPropertyName("chartType")]
        public string? ChartType { get; set; }

        [JsonPropertyName("plotArea")]
        public PlotAreaDefinition? PlotArea { get; set; }

        [JsonPropertyName("series")]
        public List<SeriesDefinition>? Series { get; set; }

        /// <summary>col | bar</summary>
        [JsonPropertyName("barDir")]
        public string? BarDir { get; set; }

        [JsonPropertyName("axes")]
        public AxesDefinition? Axes { get; set; }

        [JsonPropertyName("legend")]
        public LegendDefinition? Legend { get; set; }

        [JsonPropertyName("dataLabels")]
        public DataLabelsDefinition? DataLabels { get; set; }

        // ── Table properties ──────────────────────────────────────────────

        [JsonPropertyName("columns")]
        public List<ColumnDefinition>? Columns { get; set; }

        [JsonPropertyName("rows")]
        public List<RowDefinition>? Rows { get; set; }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Geometry / position
    // ─────────────────────────────────────────────────────────────────────────

    public class PositionDefinition
    {
        [JsonPropertyName("x")] public long X { get; set; }
        [JsonPropertyName("y")] public long Y { get; set; }
        [JsonPropertyName("cx")] public long Cx { get; set; }
        [JsonPropertyName("cy")] public long Cy { get; set; }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Fill / border / line
    // ─────────────────────────────────────────────────────────────────────────

    public class BackgroundDefinition
    {
        [JsonPropertyName("fill")]
        public FillDefinition? Fill { get; set; }
    }

    public class FillDefinition
    {
        /// <summary>solid | none | gradient | pattern</summary>
        [JsonPropertyName("type")]
        public string? Type { get; set; }

        /// <summary>6-char hex, no #</summary>
        [JsonPropertyName("color")]
        public string? Color { get; set; }
    }

    public class BorderDefinition
    {
        /// <summary>solid | none</summary>
        [JsonPropertyName("type")]
        public string? Type { get; set; }

        [JsonPropertyName("color")]
        public string? Color { get; set; }

        /// <summary>EMU — 9525 = 0.75pt, 12700 = 1pt, 19050 = 1.5pt</summary>
        [JsonPropertyName("width")]
        public int Width { get; set; }
    }

    public class LineDefinition
    {
        [JsonPropertyName("color")]
        public string? Color { get; set; }

        /// <summary>EMU line width</summary>
        [JsonPropertyName("width")]
        public int Width { get; set; }
    }

    public class ArrowEndDefinition
    {
        /// <summary>arrow | stealth | diamond | oval | block | none</summary>
        [JsonPropertyName("type")]
        public string? Type { get; set; }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Text
    // ─────────────────────────────────────────────────────────────────────────

    public class TextDefinition
    {
        [JsonPropertyName("type")]
        public string? Type { get; set; }   // "none" to suppress

        [JsonPropertyName("body")]
        public TextBodyDefinition? Body { get; set; }
    }

    public class TextBodyDefinition
    {
        /// <summary>t | ctr | b</summary>
        [JsonPropertyName("anchor")]
        public string? Anchor { get; set; }

        [JsonPropertyName("paragraphs")]
        public List<ParagraphDefinition>? Paragraphs { get; set; }

        [JsonPropertyName("autofit")]
        public bool Autofit { get; set; }
    }

    public class ParagraphDefinition
    {
        /// <summary>left | ctr | right</summary>
        [JsonPropertyName("alignment")]
        public string? Alignment { get; set; }

        [JsonPropertyName("runs")]
        public List<RunDefinition>? Runs { get; set; }

        [JsonPropertyName("lineSpacing")]
        public int LineSpacing { get; set; }
    }

    public class RunDefinition
    {
        [JsonPropertyName("text")]
        public string? Text { get; set; }

        [JsonPropertyName("bold")]
        public bool Bold { get; set; }

        [JsonPropertyName("italic")]
        public bool Italic { get; set; }

        /// <summary>Half-points. 800 = 8pt, 1800 = 18pt, 3600 = 36pt</summary>
        [JsonPropertyName("fontSize")]
        public int FontSize { get; set; }

        [JsonPropertyName("fontFace")]
        public string? FontFace { get; set; }

        /// <summary>6-char hex</summary>
        [JsonPropertyName("color")]
        public string? Color { get; set; }

        /// <summary>Superscript/subscript in thousandths of a percent. 30000 = superscript</summary>
        [JsonPropertyName("baseline")]
        public int Baseline { get; set; }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Chart
    // ─────────────────────────────────────────────────────────────────────────

    public class PlotAreaDefinition
    {
        [JsonPropertyName("fill")]
        public FillDefinition? Fill { get; set; }

        [JsonPropertyName("border")]
        public BorderDefinition? Border { get; set; }
    }

    public class SeriesDefinition
    {
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("color")]
        public string? Color { get; set; }

        [JsonPropertyName("negativeColor")]
        public string? NegativeColor { get; set; }

        [JsonPropertyName("smooth")]
        public bool Smooth { get; set; }

        [JsonPropertyName("markerSize")]
        public int MarkerSize { get; set; }

        [JsonPropertyName("markerColor")]
        public string? MarkerColor { get; set; }

        [JsonPropertyName("points")]
        public List<DataPoint>? Points { get; set; }
    }

    public class DataPoint
    {
        [JsonPropertyName("label")]
        public string? Label { get; set; }

        [JsonPropertyName("value")]
        public double Value { get; set; }
    }

    public class AxesDefinition
    {
        [JsonPropertyName("catAx")]
        public AxisDefinition? CatAx { get; set; }

        [JsonPropertyName("valAx")]
        public AxisDefinition? ValAx { get; set; }
    }

    public class AxisDefinition
    {
        [JsonPropertyName("visible")]
        public bool Visible { get; set; }

        [JsonPropertyName("labelColor")]
        public string? LabelColor { get; set; }

        /// <summary>Half-points</summary>
        [JsonPropertyName("labelFontSize")]
        public int LabelFontSize { get; set; }

        [JsonPropertyName("min")]
        public double? Min { get; set; }

        [JsonPropertyName("max")]
        public double? Max { get; set; }

        [JsonPropertyName("majorUnit")]
        public double? MajorUnit { get; set; }

        [JsonPropertyName("numFmt")]
        public string? NumFmt { get; set; }

        [JsonPropertyName("tickMark")]
        public string? TickMark { get; set; }

        [JsonPropertyName("axLine")]
        public BorderDefinition? AxLine { get; set; }

        [JsonPropertyName("gridLine")]
        public GridLineDefinition? GridLine { get; set; }
    }

    public class GridLineDefinition
    {
        /// <summary>none | solid</summary>
        [JsonPropertyName("type")]
        public string? Type { get; set; }

        [JsonPropertyName("color")]
        public string? Color { get; set; }
    }

    public class LegendDefinition
    {
        [JsonPropertyName("visible")]
        public bool Visible { get; set; }

        /// <summary>b | t | l | r | tr</summary>
        [JsonPropertyName("position")]
        public string? Position { get; set; }
    }

    public class DataLabelsDefinition
    {
        [JsonPropertyName("visible")]
        public bool Visible { get; set; }

        [JsonPropertyName("position")]
        public string? Position { get; set; }

        [JsonPropertyName("color")]
        public string? Color { get; set; }

        [JsonPropertyName("fontSize")]
        public int FontSize { get; set; }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Table
    // ─────────────────────────────────────────────────────────────────────────

    public class ColumnDefinition
    {
        /// <summary>EMU</summary>
        [JsonPropertyName("width")]
        public long Width { get; set; }
    }

    public class RowDefinition
    {
        /// <summary>EMU</summary>
        [JsonPropertyName("height")]
        public long Height { get; set; }

        [JsonPropertyName("cells")]
        public List<CellDefinition>? Cells { get; set; }
    }

    public class CellDefinition
    {
        [JsonPropertyName("text")]
        public string? Text { get; set; }

        [JsonPropertyName("bold")]
        public bool Bold { get; set; }

        [JsonPropertyName("italic")]
        public bool Italic { get; set; }

        [JsonPropertyName("fontSize")]
        public int FontSize { get; set; }

        [JsonPropertyName("color")]
        public string? Color { get; set; }

        [JsonPropertyName("fill")]
        public FillDefinition? Fill { get; set; }

        /// <summary>left | ctr | right</summary>
        [JsonPropertyName("alignment")]
        public string? Alignment { get; set; }
    }
}