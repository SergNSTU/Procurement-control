using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Страница «Заметки». Переносит поведение блока заметок из Show-MainFormV2
/// (RRFQComparer.ps1, строки ~3138-3258): список файлов, выбор, редактирование,
/// сохранение, создание новой заметки и обновление списка.
/// </summary>
public partial class NotesViewModel : ObservableObject
{
    private readonly NotesService _notes;

    // Аналог $script:NotesLoading — защита от рекурсивной загрузки при смене выбора.
    private bool _isLoading;

    public ObservableCollection<NoteFile> Notes { get; } = new();

    [ObservableProperty]
    private NoteFile? _selectedNote;

    [ObservableProperty]
    private string _editorText = string.Empty;

    [ObservableProperty]
    private string _currentPath = string.Empty;

    [ObservableProperty]
    private string _statusMessage = string.Empty;

    public NotesViewModel(NotesService notes)
    {
        _notes = notes;
        Reload();
    }

    partial void OnSelectedNoteChanged(NoteFile? value)
    {
        if (_isLoading || value is null)
        {
            return;
        }
        LoadSelectedToEditor();
    }

    [RelayCommand]
    private void Reload()
    {
        _isLoading = true;
        try
        {
            var preferredPath = SelectedNote?.Path ?? CurrentPath;

            Notes.Clear();
            foreach (var note in _notes.GetNoteFiles())
            {
                Notes.Add(note);
            }

            var target = Notes.FirstOrDefault(n =>
                string.Equals(n.Path, preferredPath, StringComparison.OrdinalIgnoreCase))
                ?? Notes.FirstOrDefault();

            SelectedNote = target;
            if (target is not null)
            {
                LoadSelectedToEditor();
            }
            StatusMessage = Notes.Count == 0 ? "Заметок нет." : $"Заметок: {Notes.Count}";
        }
        finally
        {
            _isLoading = false;
        }
    }

    [RelayCommand]
    private void Save()
    {
        if (SelectedNote is null)
        {
            StatusMessage = "Нет выбранной заметки для сохранения.";
            return;
        }
        _notes.Save(SelectedNote.Path, EditorText);
        CurrentPath = SelectedNote.Path;
        StatusMessage = "Сохранено.";
    }

    [RelayCommand]
    private void NewNote()
    {
        var path = _notes.CreateNewNote();
        _isLoading = true;
        try
        {
            Notes.Clear();
            foreach (var note in _notes.GetNoteFiles())
            {
                Notes.Add(note);
            }
            SelectedNote = Notes.FirstOrDefault(n =>
                string.Equals(n.Path, path, StringComparison.OrdinalIgnoreCase));
            if (SelectedNote is not null)
            {
                LoadSelectedToEditor();
            }
            StatusMessage = "Создана новая заметка.";
        }
        finally
        {
            _isLoading = false;
        }
    }

    private void LoadSelectedToEditor()
    {
        if (SelectedNote is null)
        {
            EditorText = string.Empty;
            CurrentPath = string.Empty;
            return;
        }
        CurrentPath = SelectedNote.Path;
        EditorText = _notes.Read(SelectedNote.Path);
    }
}
