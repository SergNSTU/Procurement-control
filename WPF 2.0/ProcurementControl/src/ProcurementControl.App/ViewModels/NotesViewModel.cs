using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Services;
using ProcurementControl.Views;

namespace ProcurementControl.ViewModels;

public partial class NotesViewModel : ObservableObject
{
    private readonly NotesService _notes;
    private readonly DispatcherTimer _autoSaveTimer;
    private bool _isLoading;
    private bool _isDirty;
    private NoteFile? _loadedNote;

    public ObservableCollection<NoteFile> Notes { get; } = new();

    [ObservableProperty]
    private NoteFile? _selectedNote;

    [ObservableProperty]
    private string _editorText = string.Empty;

    [ObservableProperty]
    private string _statusMessage = string.Empty;

    public string NoteTitle => SelectedNote?.Name ?? "Заметки";

    public NotesViewModel(NotesService notes)
    {
        _notes = notes;
        _autoSaveTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(700) };
        _autoSaveTimer.Tick += (_, _) => SaveLoadedNote();
        Reload();
    }

    partial void OnSelectedNoteChanged(NoteFile? value)
    {
        OnPropertyChanged(nameof(NoteTitle));
        if (_isLoading)
        {
            return;
        }

        SaveLoadedNote();
        LoadNote(value);
    }

    partial void OnEditorTextChanged(string value)
    {
        if (_isLoading || _loadedNote is null)
        {
            return;
        }

        _isDirty = true;
        StatusMessage = "Сохранение...";
        _autoSaveTimer.Stop();
        _autoSaveTimer.Start();
    }

    [RelayCommand]
    private void Reload()
    {
        SaveLoadedNote();
        RefreshNotes(SelectedNote?.Path ?? _loadedNote?.Path);
        StatusMessage = Notes.Count == 0 ? "Заметок нет." : "Заметок: " + Notes.Count;
    }

    [RelayCommand]
    private void Save() => SaveLoadedNote();

    [RelayCommand]
    private void NewNote()
    {
        var dialog = new InputDialogWindow("Новая заметка", "Название заметки")
        {
            Owner = Application.Current.MainWindow,
        };
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        try
        {
            SaveLoadedNote();
            var path = _notes.CreateNewNote(dialog.Value);
            RefreshNotes(path);
            StatusMessage = "Заметка создана.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Новая заметка");
        }
    }

    [RelayCommand]
    private void RenameNote()
    {
        if (SelectedNote is null)
        {
            return;
        }

        var dialog = new InputDialogWindow("Переименовать заметку", "Название заметки", SelectedNote.Name)
        {
            Owner = Application.Current.MainWindow,
        };
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        try
        {
            SaveLoadedNote();
            var renamed = _notes.Rename(SelectedNote.Path, dialog.Value);
            RefreshNotes(renamed.Path);
            StatusMessage = "Заметка переименована.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Переименовать заметку");
        }
    }

    [RelayCommand]
    private void DeleteNote()
    {
        if (SelectedNote is null)
        {
            return;
        }

        if (Notes.Count <= 1)
        {
            MessageBox.Show("Нельзя удалить единственную заметку.", "Удалить заметку");
            return;
        }

        var answer = MessageBox.Show(
            "Переместить заметку «" + SelectedNote.Name + "» в корзину?",
            "Удалить заметку",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);
        if (answer != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            SaveLoadedNote();
            _notes.MoveToTrash(SelectedNote);
            RefreshNotes(null);
            StatusMessage = "Заметка перемещена в корзину.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Удалить заметку");
        }
    }

    private void RefreshNotes(string? preferredPath)
    {
        _isLoading = true;
        try
        {
            Notes.Clear();
            foreach (var note in _notes.GetNoteFiles())
            {
                Notes.Add(note);
            }

            SelectedNote = Notes.FirstOrDefault(note =>
                string.Equals(note.Path, preferredPath, StringComparison.OrdinalIgnoreCase))
                ?? Notes.FirstOrDefault();
        }
        finally
        {
            _isLoading = false;
        }

        LoadNote(SelectedNote);
    }

    private void LoadNote(NoteFile? note)
    {
        _autoSaveTimer.Stop();
        _isLoading = true;
        try
        {
            _loadedNote = note;
            EditorText = note is null ? string.Empty : _notes.Read(note.Path);
            _isDirty = false;
        }
        finally
        {
            _isLoading = false;
        }
    }

    private void SaveLoadedNote()
    {
        _autoSaveTimer.Stop();
        if (!_isDirty || _loadedNote is null)
        {
            return;
        }

        try
        {
            StatusMessage = "Сохранение...";
            _notes.Save(_loadedNote.Path, EditorText);
            _isDirty = false;
            StatusMessage = "Сохранено";
        }
        catch (Exception ex)
        {
            StatusMessage = "Не удалось сохранить: " + ex.Message;
        }
    }
}
