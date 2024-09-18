using DocumentFormat.OpenXml;

namespace Converter.Writer
{
    internal class HelperTools
    {
        public static T GetOrCreateChild<T>(OpenXmlCompositeElement parent)
            where T : OpenXmlElement, new()
        {
            T child = parent.GetFirstChild<T>();

            if (child is null)
            {
                child = new T();
                parent.AppendChild(child);
            }

            return child;
        }
    }
}
