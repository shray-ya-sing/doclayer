using D = DocumentFormat.OpenXml.Drawing;


namespace OpenXMLExtensions
{
    public static class SolidFillExtensions
    {
        public static void SetSchemeFill(this D.SolidFill solidFill, int accentNum)
        {
                D.SchemeColorValues schemeColorVal;

                switch (accentNum)
                {
                    case 1:
                        schemeColorVal = D.SchemeColorValues.Accent1;
                        break;

                    case 2:
                        schemeColorVal = D.SchemeColorValues.Accent2;
                        break;

                    case 3:
                        schemeColorVal = D.SchemeColorValues.Accent3;
                        break;

                    case 4:
                        schemeColorVal = D.SchemeColorValues.Accent4;
                        break;
                }

            if (solidFill.GetFirstChild<D.SchemeColor>()!= null)
            {
                D.SchemeColor newCol = new D.SchemeColor() { Val = schemeColorVal };
                D.SchemeColor oldCol = solidFill.GetFirstChild<D.SchemeColor>();
                oldCol.Val = schemeColorVal;

            }
            else
            {
                D.SchemeColor newCol = new D.SchemeColor() { Val = schemeColorVal };
                solidFill.AddChild(newCol);
            }
        }

        public static void SetHexFill(this D.SolidFill solidFill, string rgbColorHex)
        {         

            if (solidFill.GetFirstChild<D.SchemeColor>() != null)
            {                
                D.SchemeColor oldCol = solidFill.GetFirstChild<D.SchemeColor>();
                oldCol.Remove();
                solidFill.AddChild(new D.RgbColorModelHex() { Val = rgbColorHex });
            }
            else if (solidFill.GetFirstChild<D.RgbColorModelHex>() != null)
            {
                solidFill.GetFirstChild<D.RgbColorModelHex>().Val = rgbColorHex;                
            }
            else if (solidFill.GetFirstChild<D.RgbColorModelHex>() is null)
            {
                solidFill.AddChild(new D.RgbColorModelHex() { Val = rgbColorHex });
            }
        }
    }
}
