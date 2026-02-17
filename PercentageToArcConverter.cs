using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace WinOptApp;

public sealed class PercentageToArcConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is not double percent)
        {
            return Geometry.Empty;
        }

        var p = Math.Max(0, Math.Min(100, percent));
        if (p <= 0)
        {
            return Geometry.Empty;
        }

        if (p >= 99.999)
        {
            return new EllipseGeometry(new System.Windows.Point(70, 70), 63, 63);
        }

        const double centerX = 70;
        const double centerY = 70;
        const double radius = 63;
        const double startAngle = -90;

        var sweepAngle = 359.999 * (p / 100.0);
        var endAngle = startAngle + sweepAngle;

        var start = PointOnCircle(centerX, centerY, radius, startAngle);
        var end = PointOnCircle(centerX, centerY, radius, endAngle);

        var isLargeArc = sweepAngle > 180;

        var figure = new PathFigure { StartPoint = start, IsClosed = false, IsFilled = false };
        figure.Segments.Add(new ArcSegment
        {
            Point = end,
            Size = new System.Windows.Size(radius, radius),
            IsLargeArc = isLargeArc,
            SweepDirection = SweepDirection.Clockwise
        });

        var geometry = new PathGeometry();
        geometry.Figures.Add(figure);
        return geometry;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return Binding.DoNothing;
    }

    private static System.Windows.Point PointOnCircle(double centerX, double centerY, double radius, double angleDegrees)
    {
        var angleRadians = angleDegrees * Math.PI / 180.0;
        return new System.Windows.Point(
            centerX + (radius * Math.Cos(angleRadians)),
            centerY + (radius * Math.Sin(angleRadians)));
    }
}
