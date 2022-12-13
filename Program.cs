namespace HogStatGenerator
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var fileName = args[0];
            double[] marksPercents = new double[] { Double.Parse(args[1]) / 100, Double.Parse(args[2]) / 100, Double.Parse(args[3]) / 100 };
            //var fileName = "C:\\Tmp\\testFile.xlsx";
            //double[] marksPercentTMP = new double[] { 0.5, 0.7, 0.9 };
            using (var xlsGen = new XlsGenerator(fileName, marksPercents))
            {
                xlsGen.Generate();
            }
            
        }
    }
}