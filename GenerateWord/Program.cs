
using Xceed.Words.NET;

namespace GenerateWord
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string templatePath = "template.docx"; // Путь к шаблону
            string outputPath = "output.docx";     // Путь для сохранения нового документа

            // Копируем шаблонный файл
            File.Copy(templatePath, outputPath, true);

            // Открываем скопированный файл как документ DocX
            using var document = DocX.Load(outputPath);
            // Заменяем текстовые плейсхолдеры
            document.ReplaceText("{{Name}}", "Иван Иванович");
            document.ReplaceText("{{Date}}", DateTime.Now.ToString("dd.MM.yyyy"));

            // Сохраняем изменения
            document.Save();
        }
    }
}
