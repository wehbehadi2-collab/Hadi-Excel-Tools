
namespace Hadi.Excel
{
    public class MappedJunction
    {
        public string Label { get; private set; }

        public double Elevation { get; private set; }

        public double X { get; private set; }

        public double Y { get; private set; }

        public bool CoordinatesGenerated { get; private set; }

        public MappedJunction(
            string label,
            double elevation,
            double x,
            double y,
            bool coordinatesGenerated)
        {
            Label = label;
            Elevation = elevation;
            X = x;
            Y = y;
            CoordinatesGenerated = coordinatesGenerated;
        }
    }
}
