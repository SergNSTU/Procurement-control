using System.IO;
using System.Text;

namespace ProcurementControl.Services;

/// <summary>Одна заметка: отображаемое имя и полный путь к файлу.</summary>
public sealed record NoteFile(string Name, string Path)
{
    public override string ToString() => Name;
}

/// <summary>
/// Работа с заметками. Переносит логику из app/modules/PurchaseStore.ps1:
/// Get-ProjectNoteFiles, Get-NoteDisplayName, Get-NotesFilePath,
/// Read-ProjectNotes, Save-ProjectNotes, New-ProjectNoteFile.
/// Заметки хранятся в {AppRoot}/data/notes как *.txt (общая папка с оригиналом).
/// </summary>
public sealed class NotesService
{
    // Оригинальный дефолтный файл: data/notes/notes.txt (Get-DefaultNotesFilePath).
    private const string DefaultNoteFileName = "notes.txt";
    // Файл, создаваемый при пустой папке (Get-ProjectNoteFiles).
    private const string SeedNoteFileName = "Общие инструкции.txt";

    /// <summary>Возвращает список *.txt в папке заметок, отсортированный по имени.
    /// Если папка пуста — создаёт стартовую заметку (как в оригинале).</summary>
    public IReadOnlyList<NoteFile> GetNoteFiles()
    {
        var root = AppPaths.NotesDirectory;
        Directory.CreateDirectory(root);

        var files = Directory.EnumerateFiles(root, "*.txt", SearchOption.TopDirectoryOnly)
            .OrderBy(p => Path.GetFileNameWithoutExtension(p), StringComparer.OrdinalIgnoreCase)
            .Select(p => new NoteFile(GetDisplayName(p), p))
            .ToList();

        if (files.Count == 0)
        {
            var seedPath = Path.Combine(root, SeedNoteFileName);
            File.WriteAllText(seedPath, string.Empty, new UTF8Encoding(false));
            files.Add(new NoteFile(GetDisplayName(seedPath), seedPath));
        }

        return files;
    }

    /// <summary>Отображаемое имя = имя файла без расширения (Get-NoteDisplayName).</summary>
    public static string GetDisplayName(string path)
        => string.IsNullOrWhiteSpace(path) ? string.Empty : Path.GetFileNameWithoutExtension(path);

    /// <summary>Читает содержимое заметки (Read-ProjectNotes).</summary>
    public string Read(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return string.Empty;
        }
        return File.ReadAllText(path, Encoding.UTF8);
    }

    /// <summary>Сохраняет текст заметки (Save-ProjectNotes).</summary>
    public void Save(string path, string text)
    {
        Directory.CreateDirectory(AppPaths.NotesDirectory);
        var target = string.IsNullOrWhiteSpace(path) ? GetDefaultNotePath() : path;
        File.WriteAllText(target, text ?? string.Empty, new UTF8Encoding(false));
    }

    public NoteFile Rename(string path, string title)
    {
        var source = ValidateNotePath(path);
        var target = Path.Combine(AppPaths.NotesDirectory, GetSafeNoteFileName(title));
        if (string.Equals(source, target, StringComparison.OrdinalIgnoreCase))
        {
            return new NoteFile(GetDisplayName(source), source);
        }

        if (File.Exists(target))
        {
            throw new InvalidOperationException("Заметка с таким названием уже существует.");
        }

        File.Move(source, target);
        return new NoteFile(GetDisplayName(target), target);
    }

    public void MoveToTrash(NoteFile note)
    {
        var path = ValidateNotePath(note.Path);
        PurchaseWriteRepository.MoveFileToTrash("note", "Заметка: " + note.Name, path);
    }

    /// <summary>Создаёт новую заметку и возвращает её путь (New-ProjectNoteFile).</summary>
    public string CreateNewNote(string? title = null)
    {
        var root = AppPaths.NotesDirectory;
        Directory.CreateDirectory(root);

        var fileName = GetSafeNoteFileName(title);
        var stem = Path.GetFileNameWithoutExtension(fileName);
        var extension = Path.GetExtension(fileName);

        var path = Path.Combine(root, fileName);
        for (int index = 2; File.Exists(path); index++)
        {
            path = Path.Combine(root, $"{stem} ({index}){extension}");
        }

        File.WriteAllText(path, string.Empty, new UTF8Encoding(false));
        return path;
    }

    public string GetDefaultNotePath() => Path.Combine(AppPaths.NotesDirectory, DefaultNoteFileName);

    private static string ValidateNotePath(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            throw new InvalidOperationException("Файл заметки не найден.");
        }

        var root = Path.GetFullPath(AppPaths.NotesDirectory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        var fullPath = Path.GetFullPath(path);
        if (!fullPath.StartsWith(root, StringComparison.OrdinalIgnoreCase)
            || !string.Equals(Path.GetExtension(fullPath), ".txt", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Можно работать только с заметками из папки приложения.");
        }

        return fullPath;
    }

    /// <summary>Санитизация имени файла (Get-SafeNoteFileName).</summary>
    private static string GetSafeNoteFileName(string? title)
    {
        var name = string.IsNullOrWhiteSpace(title) ? "Новая заметка" : title.Trim();
        foreach (var ch in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(ch.ToString(), "_");
        }
        name = System.Text.RegularExpressions.Regex.Replace(name, @"\s+", " ").Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            name = "Новая заметка";
        }
        if (!name.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
        {
            name += ".txt";
        }
        return name;
    }
}
