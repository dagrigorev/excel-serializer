namespace DemoApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var objects = ExcelSerializer.ExcelSerializer.DeserializeFromExcel<List<SampleClass>>("SampleClassValues.xlsx");

            Console.WriteLine($"Deserialized: {string.Join(';', objects.Select(o => o.ToString()))}\n");
            Console.ReadKey();
        }
    }
}
