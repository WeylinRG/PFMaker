using PFMakerLib;

namespace PFMakerConsole
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Specify the path to json config file.");
                return;
            }

            string inputPath = args[0];

            string jsonConfigPath = Path.IsPathRooted(inputPath)
                ? inputPath
                : Path.Combine(Directory.GetCurrentDirectory(), inputPath);


            Console.WriteLine("Creating print form...");
            Console.WriteLine($"JSON Config: {jsonConfigPath}");
            var pfmaker = new PFMaker();
            //string result = pfmaker.MakePrintForm(@"d:\temp\pftest\bnkros360_compV4pr_Expert.json");
            var result = pfmaker.MakePrintForm(jsonConfigPath);

            if (result.IsSuccess)
            {
                Console.WriteLine($"Done!");
            }
            else
            {
                Console.WriteLine($"Error: {result.Error}");
            }
        }
    }
}