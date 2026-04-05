using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Npgsql;

class Program
{
    static string folder = @"target folder";
    static string connStr = "Host=localhost;Database=test;Username=postgres;Password=123456";

    static void Main()
    {
        Directory.CreateDirectory(folder);
        while (true)
        {
            Console.Clear();
            Console.WriteLine("=== SIMPLE FILE TOOL ===");
            Console.WriteLine("1.Word  2.JSON/TXT  3.PG  0.Выход");
            string choice = Console.ReadLine();
            switch (choice) { case "1": WordMenu(); break; case "2": JsonMenu(); break; case "3": PostgresMenu(); break; case "0": return; }
        }
    }

    static void WordMenu()
    {
        Console.Clear();
        var docs = Directory.GetFiles(folder, "*.docx");
        if (docs.Length == 0) { Console.WriteLine("Нет .docx!"); Console.ReadKey(); return; }
        Console.WriteLine("DOCX в " + folder);
        for (int i = 0; i < docs.Length; i++) Console.WriteLine($"{i + 1}. {Path.GetFileName(docs[i])}");

        Console.Write("Номер (0=back): ");
        if (!int.TryParse(Console.ReadLine(), out int n) || n == 0 || n > docs.Length) { Console.ReadKey(); return; }
        string path = docs[n - 1];

        Console.WriteLine("1.View 2.Read 3.Edit");
        string act = Console.ReadLine();
        try
        {
            bool isEdit = act == "3";
            using var doc = WordprocessingDocument.Open(path, isEdit);

            var mainPart = doc.MainDocumentPart;
            if (mainPart == null)
            {
                mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
            }
            var document = mainPart.Document;
            if (document == null)
            {
                document = new Document(new Body());
                mainPart.Document = document;
            }
            var body = document.Body ?? new Body();
            document.Body = body;

            if (act == "1")
                Console.WriteLine(body.InnerText ?? "Пусто");
            else if (act == "2")
            {
                StringBuilder sb = new();
                foreach (var textEl in body.Descendants<Text>())
                    sb.Append(textEl.Text);
                Console.WriteLine(sb.ToString() ?? "Пусто");
            }
            else // Edit
            {
                Console.Write("Добавить текст: ");
                string newText = Console.ReadLine();

                var newPara = new Paragraph(new Run(new Text(newText)));
                body.Append(newPara);

                mainPart.Document.Save();
                Console.WriteLine("Добавлено!");
            }
        }
        catch (Exception ex) { Console.WriteLine($"ERR: {ex.Message}"); }
        Console.ReadKey();
    }

    static void JsonMenu()
    {
        Console.Clear();
        var files = Directory.GetFiles(folder, "*.json").Concat(Directory.GetFiles(folder, "*.txt")).ToArray();
        if (files.Length == 0) { Console.WriteLine("Нет файлов!"); Console.ReadKey(); return; }
        Console.WriteLine("JSON/TXT в " + folder);
        for (int i = 0; i < files.Length; i++) Console.WriteLine($"{i + 1}. {Path.GetFileName(files[i])}");

        Console.Write("Номер (0=back): ");
        if (!int.TryParse(Console.ReadLine(), out int n) || n == 0 || n > files.Length) { Console.ReadKey(); return; }
        string path = files[n - 1];
        string content = File.ReadAllText(path);

        Console.WriteLine("1.View 2.Pretty 3.Edit");
        string act = Console.ReadLine();
        if (act == "1") Console.WriteLine(content);
        else if (act == "2")
        {
            try
            {
                var doc = JsonDocument.Parse(content);
                Console.WriteLine(JsonSerializer.Serialize(doc, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch { Console.WriteLine("Не JSON"); }
        }
        else if (act == "3")
        {
            Console.Write("Новый текст: ");
            File.WriteAllText(path, Console.ReadLine());
            Console.WriteLine("OK!");
        }
        Console.ReadKey();
    }

    static void PostgresMenu()
    {
        Console.Clear();
        Console.WriteLine("PG: 1.View ID  2.Upload файл из папки  3.Edit ID  0.Back");
        string act = Console.ReadLine();
        try
        {
            using var conn = new NpgsqlConnection(connStr);
            conn.Open();

            if (act == "1")
            {
                Console.Write("ID: ");
                if (int.TryParse(Console.ReadLine(), out int id))
                {
                    using var cmd = new NpgsqlCommand("SELECT content, name FROM files WHERE id=@id", conn);
                    cmd.Parameters.AddWithValue("id", id);
                    using var reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        byte[] bytes = (byte[])reader["content"];
                        string name = reader["name"].ToString();

                        Console.WriteLine($"Файл: {name} ({bytes.Length} байт)");

                        if (name.EndsWith(".txt") || name.EndsWith(".json"))
                        {
                            Console.WriteLine("ТЕКСТ:\n" + Encoding.UTF8.GetString(bytes));
                        }
                        else if (name.EndsWith(".docx"))
                        {
                            Console.WriteLine("[DOCX] Бинарный файл. Сохрани для просмотра в Word.");
                            Console.Write("Сохранить как: ");
                            string outPath = Console.ReadLine();
                            if (!string.IsNullOrEmpty(outPath))
                            {
                                File.WriteAllBytes(outPath, bytes);
                                Console.WriteLine("✓ Сохранено!");
                            }
                        }
                        else
                        {
                            Console.WriteLine("[BIN] Первые 200 байт:\n" + Encoding.UTF8.GetString(bytes.Take(200).ToArray()));
                        }
                    }
                    else Console.WriteLine("Не найдено");
                }
            }
            else if (act == "2")
            {
                var files = Directory.GetFiles(folder);
                Console.WriteLine("Файлы для upload:");
                for (int i = 0; i < files.Length; i++) Console.WriteLine($"{i + 1}. {Path.GetFileName(files[i])}");
                Console.Write("Номер: ");
                if (int.TryParse(Console.ReadLine(), out int n) && n > 0 && n <= files.Length)
                {
                    string filePath = files[n - 1];
                    byte[] data = File.ReadAllBytes(filePath);
                    Console.Write("Имя в БД: ");
                    string name = Console.ReadLine();
                    using var cmd = new NpgsqlCommand("INSERT INTO files (type, name, content) VALUES ('file', @name, @data) RETURNING id", conn);
                    cmd.Parameters.AddWithValue("name", name);
                    cmd.Parameters.AddWithValue("data", data);
                    int newId = (int)cmd.ExecuteScalar();
                    Console.WriteLine($"Загружен ID={newId}!");
                }
            }
            else if (act == "3")
            {
                Console.Write("ID: ");
                if (int.TryParse(Console.ReadLine(), out int id))
                {
                    Console.Write("Новый текст: ");
                    byte[] data = Encoding.UTF8.GetBytes(Console.ReadLine());
                    using var cmd = new NpgsqlCommand("UPDATE files SET content=@data WHERE id=@id", conn);
                    cmd.Parameters.AddWithValue("data", data);
                    cmd.Parameters.AddWithValue("id", id);
                    Console.WriteLine(cmd.ExecuteNonQuery() > 0 ? "Обновлено!" : "Не найдено");
                }
            }
        }
        catch (Exception ex) { Console.WriteLine($"БД: {ex.Message}"); }
        Console.ReadKey();
    }
}
