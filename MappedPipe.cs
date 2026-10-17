
namespace Hadi.Excel
{
    public class MappedPipe
    {
        public string PipeId { get; private set; }
        public string StartNode { get; private set; }
        public string EndNode { get; private set; }
        public double Length { get; private set; }
        public double Diameter { get; private set; }

        public MappedPipe(
            string pipeId,
            string startNode,
            string endNode,
            double length,
            double diameter)
        {
            PipeId = pipeId;
            StartNode = startNode;
            EndNode = endNode;
            Length = length;
            Diameter = diameter;
        }
    }
}
