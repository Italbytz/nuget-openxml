using DocumentFormat.OpenXml.Presentation;

namespace Italbytz.OpenXml;

public static class ShapeExtensions
{
    public static bool IsHeader(this Shape shape)
    {
        var placeholderShape = shape.NonVisualShapeProperties
            ?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();

        if (placeholderShape is not null && placeholderShape.Type is not null &&
            placeholderShape.Type.HasValue)
            return placeholderShape.Type == PlaceholderValues.Header;

        return false;
    }

    public static bool IsFooter(this Shape shape)
    {
        var placeholderShape = shape.NonVisualShapeProperties
            ?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();

        if (placeholderShape is not null && placeholderShape.Type is not null &&
            placeholderShape.Type.HasValue)
            return placeholderShape.Type == PlaceholderValues.Footer ||
                   placeholderShape.Type == PlaceholderValues.SlideNumber ||
                   placeholderShape.Type == PlaceholderValues.DateAndTime;

        return false;
    }

    public static bool IsTitle(this Shape shape)
    {
        var placeholderShape = shape.NonVisualShapeProperties
            ?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();

        if (placeholderShape is not null && placeholderShape.Type is not null &&
            placeholderShape.Type.HasValue)
            return placeholderShape.Type == PlaceholderValues.Title ||
                   placeholderShape.Type == PlaceholderValues.CenteredTitle;

        return false;
    }
}