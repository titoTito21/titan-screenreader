using System.Windows.Forms;

namespace ScreenReader;

class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Console.WriteLine("Aplikacja Czytnika Ekranu");
        Console.WriteLine("=========================");
        Console.WriteLine();

        using (var engine = new ScreenReaderEngine())
        {
            // Set up Ctrl+C handler for graceful shutdown
            Console.CancelKeyPress += (sender, e) =>
            {
                e.Cancel = true;
                Console.WriteLine("\nZamykanie...");
                engine.Stop();
                Application.Exit();
            };

            engine.Start();

            // Run Windows Forms message loop to keep the application alive
            Application.Run();
        }
    }
}
