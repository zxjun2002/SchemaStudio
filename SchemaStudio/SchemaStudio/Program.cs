using System.Text;

namespace SchemaStudio
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); // ExcelDataReader 需要
            ApplicationConfiguration.Initialize();
            Application.Run(new UI.MainForm());
        }
    }
}