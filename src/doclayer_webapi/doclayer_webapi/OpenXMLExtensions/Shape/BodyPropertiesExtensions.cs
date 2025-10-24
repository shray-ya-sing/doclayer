using D= DocumentFormat.OpenXml.Drawing;


namespace OpenXMLExtensions
{
    public static class BodyPropertiesExtensions
    {
        /// <summary>
        /// Sets shape autofit
        /// </summary>
        /// <param name="bodyProperties"></param>
        public static void SetShapeAutofit(this D.BodyProperties bodyProperties)
        {
            if (bodyProperties.GetFirstChild<D.NoAutoFit>() != null)
            {

                bodyProperties.RemoveChild<D.NoAutoFit>(bodyProperties.GetFirstChild<D.NoAutoFit>());
            }
            if (bodyProperties.GetFirstChild<D.NormalAutoFit>() != null)
            {

                bodyProperties.RemoveChild<D.NormalAutoFit>(bodyProperties.GetFirstChild<D.NormalAutoFit>());
            }

            if (bodyProperties.GetFirstChild<D.ShapeAutoFit>() is null)
            {
                bodyProperties.AddChild(new D.ShapeAutoFit());
            }
            else
            {
                return;               

            }
        }

        /// <summary>
        /// Sets no auto fit
        /// </summary>
        /// <param name="bodyProperties"></param>
        public static void SetNoAutofit(this D.BodyProperties bodyProperties)
        {
            if (bodyProperties.GetFirstChild<D.NoAutoFit>() is null)
            {
                if (bodyProperties.GetFirstChild<D.ShapeAutoFit>() != null)
                {
                    
                    bodyProperties.RemoveChild<D.ShapeAutoFit>(bodyProperties.GetFirstChild<D.ShapeAutoFit>());
                }
                if (bodyProperties.GetFirstChild<D.NormalAutoFit>() != null)
                {

                    bodyProperties.RemoveChild<D.NormalAutoFit>(bodyProperties.GetFirstChild<D.NormalAutoFit>());
                }

                bodyProperties.AddChild(new D.NoAutoFit());
            }
            else
            {
                return;

            }
        }
    }
}
