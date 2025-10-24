using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;

namespace OpenXMLExtensions
{
    public static class ShapeStyleExtensions
    {
        public static void SetDefaultReferences( this P.ShapeStyle shapeStyle)
        {
            D.SchemeColor lineColor = new D.SchemeColor() { Val = SchemeColorValues.Accent1 };

            D.LineReference lineReference = new D.LineReference() { Index = 2 };
            lineReference.AddChild(lineColor);
            shapeStyle.AddChild(lineReference);

            D.SchemeColor schemeColor = new D.SchemeColor() { Val = SchemeColorValues.Accent1 };
            D.FillReference fillReference = new D.FillReference() { Index = 1 };
            fillReference.AddChild(schemeColor);
            shapeStyle.AddChild(fillReference);

            D.SchemeColor effectColor=new D.SchemeColor() { Val = SchemeColorValues.Accent1 };
            D.EffectReference effectReference = new D.EffectReference() { Index = 0 };
            effectReference.AddChild(effectColor);
            shapeStyle.AddChild(effectReference);

            D.FontReference fontReference = new D.FontReference() { Index = FontCollectionIndexValues.Minor };
            D.SchemeColor fontColor = new D.SchemeColor() { Val = SchemeColorValues.Light1 };
            fontReference.AddChild(fontColor);
            shapeStyle.AddChild(fontReference);
        }
    }
}