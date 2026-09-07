using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;
using ProcurementControl.Models;
using System.Text.Json;

namespace ProcurementControl.Services;

/// <summary>
/// Write-операции WPF: порты функций записи оригинала (настройки, напоминания,
/// база квот, корзина, уведомления, поступления). Первые безопасные операции для
/// deals/deal_suppliers — порты New-PurchaseDeal, Update-PurchaseDeal,
/// Set-PurchaseDealArchived и Add-PurchaseSupplier. Удаление всегда идёт через
/// корзину (таблица trash_items), как в New-PurchaseTrashItem.
/// </summary>
public static class PurchaseWriteRepository
{
    private const string OrderedStatus = "Заказано";

    /// <summary>Добавляет WPF-колонку суммы задачи, не изменяя код WinForms.</summary>
    public static void EnsureComponentOrderAmountColumn()
    {
        PurchaseWriteSession.Execute(connection =>
        {
            using var check = connection.CreateCommand();
            check.CommandText = "PRAGMA table_info(component_deals);";
            using var reader = check.ExecuteReader();
            var exists = false;
            while (reader.Read())
            {
                if (string.Equals(reader.GetString(1), "order_amount", StringComparison.OrdinalIgnoreCase))
                {
                    exists = true;
                    break;
                }
            }

            if (!exists)
            {
                using var alter = connection.CreateCommand();
                alter.CommandText = "ALTER TABLE component_deals ADD COLUMN order_amount TEXT;";
                alter.ExecuteNonQuery();
            }
            return 0;
        });
    }

    public static long CreateOrUpdateDeal(PurchaseDealDraft draft)
    {
        ArgumentNullException.ThrowIfNull(draft);
        var dealNumber = RequireDealNumber(draft.DealNumber);
        var status = NormalizeStatus(draft.Status);

        return PurchaseWriteSession.Execute(connection =>
        {
            var now = NowText();
            using var command = connection.CreateCommand();
            command.CommandText = @"
INSERT INTO deals(
    deal_number, client, status, title, comment, workflow_template_id, masks,
    created_at, ordered_at, updated_at)
VALUES(
    @deal_number, @client, @status, @title, @comment,
    (SELECT id FROM workflow_templates WHERE is_default = 1 ORDER BY id LIMIT 1),
    2, @created_at, @ordered_at, @updated_at)
ON CONFLICT(deal_number) DO UPDATE SET
    client = CASE WHEN excluded.client <> '' THEN excluded.client ELSE client END,
    status = excluded.status,
    ordered_at = CASE
        WHEN IFNULL(ordered_at, '') <> '' THEN ordered_at
        WHEN excluded.ordered_at IS NOT NULL THEN excluded.ordered_at
        ELSE ordered_at
    END,
    title = CASE WHEN excluded.title <> '' THEN excluded.title ELSE title END,
    comment = CASE WHEN excluded.comment <> '' THEN excluded.comment ELSE comment END,
    archived = 0,
    updated_at = excluded.updated_at;";
            Add(command, "@deal_number", dealNumber);
            Add(command, "@client", draft.Client.Trim());
            Add(command, "@status", status);
            Add(command, "@title", draft.Title.Trim());
            Add(command, "@comment", draft.Comment.Trim());
            Add(command, "@created_at", now);
            Add(command, "@ordered_at", status == OrderedStatus ? now : null);
            Add(command, "@updated_at", now);
            command.ExecuteNonQuery();

            var dealId = GetDealId(connection, dealNumber);
            WriteActivity(connection, "deal", dealId, "Создана/обновлена сделка", dealNumber, dealId);
            return dealId;
        });
    }

    public static void UpdateDeal(long dealId, PurchaseDealUpdate update)
    {
        if (dealId <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(dealId), "Выберите сделку.");
        }
        ArgumentNullException.ThrowIfNull(update);

        PurchaseWriteSession.Execute(connection =>
        {
            var previousStatus = GetString(connection, "SELECT IFNULL(status, '') FROM deals WHERE id = @id", dealId);
            if (previousStatus is null)
            {
                throw new InvalidOperationException("Сделка не найдена.");
            }

            var status = NormalizeStatus(update.Status);
            var dealNumber = string.IsNullOrWhiteSpace(update.DealNumber)
                ? null
                : RequireDealNumber(update.DealNumber);
            var orderedAt = previousStatus.Trim() != OrderedStatus && status == OrderedStatus ? NowText() : null;
            var boardCount = NormalizeBoardCount(update.BoardCount);

            using var command = connection.CreateCommand();
            command.CommandText = @"
UPDATE deals SET
    deal_number = CASE WHEN @deal_number IS NULL THEN deal_number ELSE @deal_number END,
    board_count = CASE WHEN @board_count IS NULL THEN board_count ELSE @board_count END,
    period = CASE WHEN @period IS NULL THEN period ELSE @period END,
    priority = CASE WHEN @priority IS NULL THEN priority ELSE @priority END,
    client = @client,
    status = @status,
    ordered_at = CASE
        WHEN IFNULL(ordered_at, '') <> '' THEN ordered_at
        WHEN @ordered_at IS NOT NULL THEN @ordered_at
        ELSE ordered_at
    END,
    tracking_status = CASE WHEN @tracking_status IS NULL THEN tracking_status ELSE @tracking_status END,
    executor = CASE WHEN @executor IS NULL THEN executor ELSE @executor END,
    reminder_date = CASE WHEN @reminder_date IS NULL THEN reminder_date ELSE @reminder_date END,
    assembly_location = CASE WHEN @assembly_location IS NULL THEN assembly_location ELSE @assembly_location END,
    comment = @comment,
    updated_at = @updated_at
WHERE id = @id;";
            Add(command, "@deal_number", dealNumber);
            Add(command, "@board_count", boardCount);
            Add(command, "@period", TrimOrNull(update.Period));
            Add(command, "@priority", TrimOrNull(update.Priority));
            Add(command, "@client", update.Client.Trim());
            Add(command, "@status", status);
            Add(command, "@ordered_at", orderedAt);
            Add(command, "@tracking_status", TrimOrNull(update.TrackingStatus));
            Add(command, "@executor", TrimOrNull(update.Executor));
            Add(command, "@reminder_date", TrimOrNull(update.ReminderDate));
            Add(command, "@assembly_location", TrimOrNull(update.AssemblyLocation));
            Add(command, "@comment", update.Comment.Trim());
            Add(command, "@updated_at", NowText());
            Add(command, "@id", dealId);
            command.ExecuteNonQuery();

            WriteActivity(connection, "deal", dealId, "Изменена сделка", "Этап: " + status, dealId);
            return 0;
        });
    }

    /// <summary>Изменяет только количество плат без перезаписи остальных полей сделки.</summary>
    public static void UpdateDealBoardCount(long dealId, string? boardCount)
    {
        if (dealId <= 0) throw new ArgumentOutOfRangeException(nameof(dealId), "Выберите сделку.");
        var normalized = NormalizeBoardCount(boardCount);
        if (normalized is null) throw new ArgumentException("Укажите количество плат.", nameof(boardCount));

        PurchaseWriteSession.Execute(connection =>
        {
            using var command = connection.CreateCommand();
            command.CommandText = @"
UPDATE deals SET board_count = @board_count, updated_at = @updated_at WHERE id = @id;";
            Add(command, "@board_count", normalized);
            Add(command, "@updated_at", NowText());
            Add(command, "@id", dealId);
            if (command.ExecuteNonQuery() == 0) throw new InvalidOperationException("Сделка не найдена.");
            WriteActivity(connection, "deal", dealId, "Изменено количество плат", normalized, dealId);
            return 0;
        });
    }

    /// <summary>Изменяет дату напоминания сделки из календаря.</summary>
    public static void UpdateDealReminderDate(long dealId, string reminderDate)
    {
        if (dealId <= 0) throw new ArgumentOutOfRangeException(nameof(dealId), "Выберите сделку.");
        PurchaseWriteSession.Execute(connection =>
        {
            using var command = connection.CreateCommand();
            command.CommandText = "UPDATE deals SET reminder_date = @reminder_date, updated_at = @updated_at WHERE id = @id;";
            Add(command, "@reminder_date", reminderDate.Trim());
            Add(command, "@updated_at", NowText());
            Add(command, "@id", dealId);
            if (command.ExecuteNonQuery() == 0) throw new InvalidOperationException("Сделка не найдена.");
            WriteActivity(connection, "deal", dealId, "Изменено напоминание", reminderDate.Trim(), dealId);
            return 0;
        });
    }

    /// <summary>Сохраняет одно редактируемое поле строки сделок, не затрагивая остальные.</summary>
    public static void UpdateDealGridField(long dealId, string field, string? value)
    {
        if (dealId <= 0) throw new ArgumentOutOfRangeException(nameof(dealId), "Выберите сделку.");

        var column = field switch
        {
            nameof(PurchaseDealRow.DealNumber) => "deal_number",
            nameof(PurchaseDealRow.BoardCount) => "board_count",
            nameof(PurchaseDealRow.Client) => "client",
            nameof(PurchaseDealRow.Priority) => "priority",
            nameof(PurchaseDealRow.Comment) => "comment",
            nameof(PurchaseDealRow.Period) => "period",
            nameof(PurchaseDealRow.Executor) => "executor",
            nameof(PurchaseDealRow.AssemblyLocation) => "assembly_location",
            nameof(PurchaseDealRow.TrackingStatus) => "tracking_status",
            nameof(PurchaseDealRow.MasksText) => "masks",
            _ => null,
        };

        PurchaseWriteSession.Execute(connection =>
        {
            var now = NowText();
            using var command = connection.CreateCommand();
            if (field == nameof(PurchaseDealRow.Status))
            {
                var previousStatus = GetString(connection, "SELECT IFNULL(status, '') FROM deals WHERE id = @id", dealId);
                if (previousStatus is null) throw new InvalidOperationException("Сделка не найдена.");
                var status = NormalizeStatus(value);
                command.CommandText = @"
UPDATE deals SET status = @value,
    ordered_at = CASE WHEN IFNULL(ordered_at, '') <> '' THEN ordered_at
                      WHEN @became_ordered = 1 THEN @updated_at ELSE ordered_at END,
    updated_at = @updated_at WHERE id = @id;";
                Add(command, "@value", status);
                Add(command, "@became_ordered", previousStatus.Trim() != OrderedStatus && status == OrderedStatus ? 1 : 0);
            }
            else if (column is not null)
            {
                command.CommandText = "UPDATE deals SET " + column + " = @value, updated_at = @updated_at WHERE id = @id;";
                Add(command, "@value", field == nameof(PurchaseDealRow.MasksText)
                    ? NormalizeMasks(value)
                    : value?.Trim() ?? string.Empty);
            }
            else
            {
                throw new ArgumentException("Поле нельзя изменить из таблицы.", nameof(field));
            }

            Add(command, "@updated_at", now);
            Add(command, "@id", dealId);
            if (command.ExecuteNonQuery() == 0) throw new InvalidOperationException("Сделка не найдена.");
            WriteActivity(connection, "deal", dealId, "Изменено поле сделки", field, dealId);
            return 0;
        });
    }

    /// <summary>Удаляет сделку, её поставщиков, документы и связи Bitrix в корзину.</summary>
    public static void DeleteDeal(long dealId)
    {
        if (dealId <= 0) throw new ArgumentOutOfRangeException(nameof(dealId), "Выберите сделку.");
        PurchaseWriteSession.Execute(connection =>
        {
            var deals = SelectRows(connection, "SELECT * FROM deals WHERE id = @id", command => Add(command, "@id", dealId));
            if (deals.Count == 0) throw new InvalidOperationException("Сделка не найдена.");
            var dealNumber = Convert.ToString(deals[0].GetValueOrDefault("deal_number")) ?? string.Empty;
            var suppliers = SelectRows(connection, "SELECT * FROM deal_suppliers WHERE deal_id = @deal_id", command => Add(command, "@deal_id", dealId));
            var documents = SelectRows(connection, "SELECT * FROM documents WHERE deal_id = @deal_id", command => Add(command, "@deal_id", dealId));
            var bitrixLinks = SelectRows(connection, "SELECT * FROM bitrix_task_links WHERE deal_id = @deal_id", command => Add(command, "@deal_id", dealId));
            var files = documents.Select(row => ResolvePortable(Convert.ToString(row.GetValueOrDefault("stored_path"))))
                .Where(File.Exists).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            InsertTrashItem(connection, "deal", dealId, "Сделка: " + dealNumber,
                new Dictionary<string, object>
                {
                    ["deals"] = deals,
                    ["deal_suppliers"] = suppliers,
                    ["documents"] = documents,
                    ["bitrix_task_links"] = bitrixLinks,
                }, files, string.Empty);

            ExecuteDelete(connection, "DELETE FROM documents WHERE deal_id = @deal_id;", dealId);
            ExecuteDelete(connection, "DELETE FROM deal_suppliers WHERE deal_id = @deal_id;", dealId);
            ExecuteDelete(connection, "DELETE FROM bitrix_task_links WHERE deal_id = @deal_id;", dealId);
            ExecuteDelete(connection, "DELETE FROM deals WHERE id = @deal_id;", dealId);
            WriteActivity(connection, "deal", dealId, "Удалена сделка", dealNumber, dealId);
            return 0;
        });
    }

    public static void SetDealArchived(long dealId, bool archived)
    {
        if (dealId <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(dealId), "Выберите сделку.");
        }

        PurchaseWriteSession.Execute(connection =>
        {
            using var command = connection.CreateCommand();
            command.CommandText = "UPDATE deals SET archived = @archived, updated_at = @updated_at WHERE id = @id;";
            Add(command, "@archived", archived ? 1 : 0);
            Add(command, "@updated_at", NowText());
            Add(command, "@id", dealId);
            if (command.ExecuteNonQuery() == 0)
            {
                throw new InvalidOperationException("Сделка не найдена.");
            }

            WriteActivity(connection, "deal", dealId,
                archived ? "Сделка отправлена в архив" : "Сделка возвращена из архива", string.Empty, dealId);
            return 0;
        });
    }

    public static void AddOrTouchSupplier(long dealId, string supplier)
    {
        if (dealId <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(dealId), "Выберите сделку.");
        }
        var normalizedSupplier = supplier?.Trim() ?? string.Empty;
        if (normalizedSupplier.Length == 0)
        {
            throw new ArgumentException("Введите поставщика.", nameof(supplier));
        }

        PurchaseWriteSession.Execute(connection =>
        {
            if (GetString(connection, "SELECT deal_number FROM deals WHERE id = @id", dealId) is null)
            {
                throw new InvalidOperationException("Сделка не найдена.");
            }

            var now = NowText();
            using var command = connection.CreateCommand();
            command.CommandText = @"
INSERT INTO deal_suppliers(deal_id, supplier, created_at, updated_at)
VALUES(@deal_id, @supplier, @created_at, @updated_at)
ON CONFLICT(deal_id, supplier) DO UPDATE SET updated_at = excluded.updated_at;";
            Add(command, "@deal_id", dealId);
            Add(command, "@supplier", normalizedSupplier);
            Add(command, "@created_at", now);
            Add(command, "@updated_at", now);
            command.ExecuteNonQuery();

            using var touch = connection.CreateCommand();
            touch.CommandText = "UPDATE deals SET updated_at = @updated_at WHERE id = @id;";
            Add(touch, "@updated_at", now);
            Add(touch, "@id", dealId);
            touch.ExecuteNonQuery();

            WriteActivity(connection, "supplier", 0, "Добавлен/обновлен поставщик", normalizedSupplier, dealId);
            return 0;
        });
    }

    /// <summary>Порт Update-PurchaseSupplier: полное сохранение карточки поставщика.</summary>
    public static void UpdateSupplier(long supplierId, PurchaseSupplierUpdate update)
    {
        ArgumentNullException.ThrowIfNull(update);
        if (supplierId <= 0) throw new ArgumentOutOfRangeException(nameof(supplierId), "Выберите поставщика.");

        PurchaseWriteSession.Execute(connection =>
        {
            var dealId = ToLong(GetString(connection, "SELECT deal_id FROM deal_suppliers WHERE id = @id", supplierId));
            if (dealId <= 0) throw new InvalidOperationException("Поставщик не найден.");

            using var command = connection.CreateCommand();
            command.CommandText = @"
UPDATE deal_suppliers SET
 invoice_received=@invoice_received, invoice_confirmed=@invoice_confirmed,
 supplier_order_created=@supplier_order_created, erp_supplier_sent=@erp_supplier_sent,
 erp_roger_sent=@erp_roger_sent, pi_amount_usd=@pi_amount_usd, pi_amount_cny=@pi_amount_cny,
 pi_amount_rub=@pi_amount_rub, paid_amount=@paid_amount, delivery_weeks=@delivery_weeks,
 payment_submitted=@payment_submitted, paid=@paid, invoice_confirmed_date=@invoice_confirmed_date,
 components_receipt_date=@components_receipt_date, actual_receipt_date=@actual_receipt_date,
 comment=@comment, updated_at=@updated_at WHERE id=@id;";
            Add(command, "@invoice_received", update.InvoiceReceived ? 1 : 0);
            Add(command, "@invoice_confirmed", update.InvoiceConfirmed ? 1 : 0);
            Add(command, "@supplier_order_created", update.SupplierOrderCreated ? 1 : 0);
            Add(command, "@erp_supplier_sent", update.ErpSupplierSent ? 1 : 0);
            Add(command, "@erp_roger_sent", update.ErpRogerSent ? 1 : 0);
            Add(command, "@pi_amount_usd", RrfqEngine.ConvertToNumber(update.PiAmountUsd));
            Add(command, "@pi_amount_cny", RrfqEngine.ConvertToNumber(update.PiAmountCny));
            Add(command, "@pi_amount_rub", RrfqEngine.ConvertToNumber(update.PiAmountRub));
            Add(command, "@paid_amount", update.PaidAmount.Trim());
            Add(command, "@delivery_weeks", string.IsNullOrWhiteSpace(update.DeliveryWeeks) ? null : update.DeliveryWeeks.Trim());
            Add(command, "@payment_submitted", update.PaymentSubmitted ? 1 : 0);
            Add(command, "@paid", update.Paid ? 1 : 0);
            Add(command, "@invoice_confirmed_date", update.InvoiceConfirmedDate.Trim());
            Add(command, "@components_receipt_date", update.ComponentsReceiptDate.Trim());
            Add(command, "@actual_receipt_date", update.ActualReceiptDate.Trim());
            Add(command, "@comment", update.Comment.Trim());
            Add(command, "@updated_at", NowText());
            Add(command, "@id", supplierId);
            command.ExecuteNonQuery();
            TouchDeal(connection, dealId);
            WriteActivity(connection, "supplier", supplierId, "Изменен поставщик",
                $"Оплата: {update.Paid}; PI: {update.InvoiceReceived}; срок: {update.DeliveryWeeks}", dealId, supplierId);
            return 0;
        });
    }

    /// <summary>Порт Delete-PurchaseSupplier: документы и поставщик перемещаются в корзину одной записью.</summary>
    public static void DeleteSupplier(long supplierId)
    {
        if (supplierId <= 0) throw new ArgumentOutOfRangeException(nameof(supplierId), "Выберите поставщика.");
        PurchaseWriteSession.Execute(connection =>
        {
            var supplierRows = SelectRows(connection, "SELECT * FROM deal_suppliers WHERE id = @id", c => Add(c, "@id", supplierId));
            if (supplierRows.Count == 0) throw new InvalidOperationException("Поставщик не найден.");
            var supplierRow = supplierRows[0];
            var dealId = ToLong(supplierRow.GetValueOrDefault("deal_id"));
            var supplier = Convert.ToString(supplierRow.GetValueOrDefault("supplier")) ?? string.Empty;
            var documents = SelectRows(connection, "SELECT * FROM documents WHERE supplier_id = @supplier_id", c => Add(c, "@supplier_id", supplierId));
            var files = documents.Select(row => ResolvePortable(Convert.ToString(row.GetValueOrDefault("stored_path"))))
                .Where(File.Exists).ToList();
            InsertTrashItem(connection, "supplier", supplierId, "Поставщик: " + supplier,
                new Dictionary<string, object> { ["deal_suppliers"] = supplierRows, ["documents"] = documents }, files, string.Empty);
            using (var deleteDocuments = connection.CreateCommand())
            {
                deleteDocuments.CommandText = "DELETE FROM documents WHERE supplier_id = @supplier_id;";
                Add(deleteDocuments, "@supplier_id", supplierId); deleteDocuments.ExecuteNonQuery();
            }
            using (var deleteSupplier = connection.CreateCommand())
            {
                deleteSupplier.CommandText = "DELETE FROM deal_suppliers WHERE id = @id;";
                Add(deleteSupplier, "@id", supplierId); deleteSupplier.ExecuteNonQuery();
            }
            TouchDeal(connection, dealId);
            WriteActivity(connection, "supplier", supplierId, "Поставщик удален", supplier, dealId, supplierId);
            return 0;
        });
    }

    /// <summary>
    /// Порт Clear-ActivityLog из QuoteHistory.ps1 (строки 240-246): журнал
    /// сначала сохраняется в корзину (таблица trash_items, ключ «activity_log»
    /// в payload_json — формат Convert-PurchaseRowsToPayload), затем удаляется.
    /// </summary>
    public static void ClearActivityLog()
    {
        PurchaseWriteSession.Execute(connection =>
        {
            var rows = new List<Dictionary<string, object?>>();
            using (var select = connection.CreateCommand())
            {
                select.CommandText = "SELECT * FROM activity_log";
                using var reader = select.ExecuteReader();
                var columnCount = reader.FieldCount;
                while (reader.Read())
                {
                    var row = new Dictionary<string, object?>(columnCount);
                    for (var i = 0; i < columnCount; i++)
                    {
                        row[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                    }
                    rows.Add(row);
                }
            }

            var payload = JsonSerializer.Serialize(
                new Dictionary<string, object> { ["activity_log"] = rows });
            using (var trash = connection.CreateCommand())
            {
                trash.CommandText = @"
INSERT INTO trash_items(entity_type, entity_id, title, payload_json, files_json, deleted_at)
VALUES(@entity_type, @entity_id, @title, @payload_json, '', @deleted_at);";
                Add(trash, "@entity_type", "activity_log");
                Add(trash, "@entity_id", 0);
                Add(trash, "@title", "Очищен журнал действий");
                Add(trash, "@payload_json", payload);
                Add(trash, "@deleted_at", NowText());
                trash.ExecuteNonQuery();
            }

            using var delete = connection.CreateCommand();
            delete.CommandText = "DELETE FROM activity_log;";
            delete.ExecuteNonQuery();
            return 0;
        });
    }

    /// <summary>
    /// Порт Set-PurchaseSetting (PurchaseStore.ps1, строки 655-662).
    /// </summary>
    public static void SetSetting(string key, string value)
    {
        ArgumentException.ThrowIfNullOrEmpty(key);

        PurchaseWriteSession.Execute(connection =>
        {
            using var command = connection.CreateCommand();
            command.CommandText = @"
INSERT INTO settings(key, value) VALUES(@key, @value)
ON CONFLICT(key) DO UPDATE SET value = excluded.value;";
            Add(command, "@key", key);
            Add(command, "@value", value ?? string.Empty);
            command.ExecuteNonQuery();
            return 0;
        });
    }

    /// <summary>
    /// Порт Set-ReminderDone (PurchaseStore.ps1, строки 985-994).
    /// </summary>
    public static void SetReminderDone(long reminderId, bool done)
    {
        if (reminderId <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(reminderId), "Выберите напоминание.");
        }

        PurchaseWriteSession.Execute(connection =>
        {
            using var command = connection.CreateCommand();
            command.CommandText = "UPDATE reminders SET status = @status, updated_at = @updated_at WHERE id = @id;";
            Add(command, "@status", done ? "Done" : "Open");
            Add(command, "@updated_at", NowText());
            Add(command, "@id", reminderId);
            command.ExecuteNonQuery();
            return 0;
        });
    }

    /// <summary>
    /// Порт Clear-PurchaseDealReminder (PurchaseStore.ps1, строки 1625-1634).
    /// </summary>
    public static void ClearDealReminder(long dealId)
    {
        if (dealId <= 0)
        {
            return;
        }

        PurchaseWriteSession.Execute(connection =>
        {
            using var command = connection.CreateCommand();
            command.CommandText = "UPDATE deals SET reminder_date = '', updated_at = @updated_at WHERE id = @id;";
            Add(command, "@updated_at", NowText());
            Add(command, "@id", dealId);
            command.ExecuteNonQuery();
            return 0;
        });
    }

    /// <summary>
    /// Порт Save-Reminder (PurchaseStore.ps1, строки 957-983).
    /// </summary>
    public static void SaveReminder(string title, string dueDate, long dealId = 0, long supplierId = 0, string source = "manual")
    {
        if (string.IsNullOrWhiteSpace(title))
        {
            throw new ArgumentException("Введите текст напоминания.", nameof(title));
        }

        PurchaseWriteSession.Execute(connection =>
        {
            var now = NowText();
            using var command = connection.CreateCommand();
            command.CommandText = @"
INSERT INTO reminders(deal_id, supplier_id, component_id, title, due_date, status, source, created_at, updated_at)
VALUES(@deal_id, @supplier_id, NULL, @title, @due_date, 'Open', @source, @created_at, @updated_at);";
            Add(command, "@deal_id", dealId > 0 ? dealId : null);
            Add(command, "@supplier_id", supplierId > 0 ? supplierId : null);
            Add(command, "@title", title.Trim());
            Add(command, "@due_date", dueDate?.Trim() ?? string.Empty);
            Add(command, "@source", source);
            Add(command, "@created_at", now);
            Add(command, "@updated_at", now);
            command.ExecuteNonQuery();

            WriteActivity(connection, "reminder", 0, "Создано напоминание", title, dealId, supplierId);
            return 0;
        });
    }

    /// <summary>
    /// Порт ветки «задача» из New-QuickCaptureRecord (QuickCaptureAgent.ps1,
    /// строки 277-292): сделка создаётся из выделенного текста с признаком
    /// source='quick_capture'. Если настройка была выключена — включается,
    /// как в оригинале (строки 254-257).
    /// </summary>
    public static long CreateQuickCaptureTask(string title, string sourceText, string dueDate, string dealNumber)
    {
        var normalizedTitle = NormalizeQuickCaptureText(title);
        if (normalizedTitle.Length == 0)
        {
            throw new ArgumentException("Введите заголовок.", nameof(title));
        }

        if (normalizedTitle.Length > 240)
        {
            normalizedTitle = normalizedTitle.Substring(0, 237) + "...";
        }

        var source = NormalizeQuickCaptureText(sourceText);
        var fallback = source.Length > 0 ? source : normalizedTitle;

        return PurchaseWriteSession.Execute(connection =>
        {
            EnableQuickCaptureIfNeeded(connection);

            var now = NowText();
            using var command = connection.CreateCommand();
            command.CommandText = @"
INSERT INTO component_deals(entry_date, deal_number, status, stage, description, next_action,
 deadline_date, priority, created_at, updated_at, source, source_text)
VALUES(@entry_date, @deal_number, 'В работе', 'Запросил поставщиков', @description, @next_action,
 @deadline_date, '3', @created_at, @updated_at, 'quick_capture', @source_text);";
            Add(command, "@entry_date", DateTime.Now.ToString("dd.MM.yyyy"));
            Add(command, "@deal_number", NormalizeQuickCaptureText(dealNumber));
            Add(command, "@description", fallback);
            Add(command, "@next_action", normalizedTitle);
            Add(command, "@deadline_date", NormalizeQuickCaptureText(dueDate));
            Add(command, "@created_at", now);
            Add(command, "@updated_at", now);
            Add(command, "@source_text", fallback);
            command.ExecuteNonQuery();

            using var idCommand = connection.CreateCommand();
            idCommand.CommandText = "SELECT id FROM component_deals ORDER BY id DESC LIMIT 1;";
            var id = Convert.ToInt64(idCommand.ExecuteScalar());

            WriteActivity(connection, "component", id, "Создана задача из выделенного текста", normalizedTitle, 0);
            return id;
        });
    }

    /// <summary>
    /// Порт Normalize-QuickCaptureText: убрать NUL-символы и пробелы по краям.
    /// </summary>
    private static string NormalizeQuickCaptureText(string? value)
        => (value ?? string.Empty).Replace("\0", string.Empty).Trim();

    /// <summary>Если быстрый захват был выключен — включить, как в оригинальном агенте.</summary>
    private static void EnableQuickCaptureIfNeeded(SqliteConnection connection)
    {
        using var read = connection.CreateCommand();
        read.CommandText = "SELECT value FROM settings WHERE key = 'quick_capture.enabled';";
        var value = read.ExecuteScalar();
        if (value is not null and not DBNull && Convert.ToString(value) == "0")
        {
            using var update = connection.CreateCommand();
            update.CommandText = "UPDATE settings SET value = '1' WHERE key = 'quick_capture.enabled';";
            update.ExecuteNonQuery();
        }
    }

    /// <summary>
    /// Порт Remove-Reminder (PurchaseStore.ps1, строки 2467-2481): напоминание
    /// сначала сохраняется в корзину, затем удаляется.
    /// </summary>
    public static void RemoveReminder(long reminderId)
    {
        if (reminderId <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(reminderId), "Некорректное напоминание.");
        }

        PurchaseWriteSession.Execute(connection =>
        {
            var rows = SelectRows(connection, "SELECT * FROM reminders WHERE id = @id",
                command => Add(command, "@id", reminderId));
            if (rows.Count == 0)
            {
                throw new InvalidOperationException("Напоминание не найдено.");
            }

            var title = Convert.ToString(rows[0].GetValueOrDefault("title")) ?? string.Empty;
            InsertTrashItem(connection, "reminder", reminderId, "Напоминание: " + title,
                new Dictionary<string, object> { ["reminders"] = rows });

            using var delete = connection.CreateCommand();
            delete.CommandText = "DELETE FROM reminders WHERE id = @id;";
            Add(delete, "@id", reminderId);
            delete.ExecuteNonQuery();

            var dealId = ToLong(rows[0].GetValueOrDefault("deal_id"));
            var supplierId = ToLong(rows[0].GetValueOrDefault("supplier_id"));
            WriteActivity(connection, "reminder", reminderId, "Удалено напоминание", title, dealId, supplierId);
            return 0;
        });
    }

    /// <summary>Удаляет автоматический пункт из списка до изменения его исходного состояния.</summary>
    public static void RemoveAutomaticReminder(long dealId, long supplierId, string title)
    {
        if (dealId <= 0 || supplierId <= 0 || string.IsNullOrWhiteSpace(title))
            throw new ArgumentException("Некорректное автоматическое напоминание.");

        PurchaseWriteSession.Execute(connection =>
        {
            var key = PurchaseLogic.AutomaticReminderKey(dealId, supplierId, title);
            using var insert = connection.CreateCommand();
            insert.CommandText = @"INSERT OR IGNORE INTO reminder_suppressions(reminder_key, title, deal_id, supplier_id, created_at)
VALUES (@key, @title, @deal_id, @supplier_id, @created_at);";
            Add(insert, "@key", key);
            Add(insert, "@title", title);
            Add(insert, "@deal_id", dealId);
            Add(insert, "@supplier_id", supplierId);
            Add(insert, "@created_at", NowText());
            insert.ExecuteNonQuery();

            InsertTrashItem(connection, "automatic_reminder", 0, "Автоматическое напоминание: " + title,
                new Dictionary<string, object> { ["reminder_suppressions"] = new List<Dictionary<string, object>>
                {
                    new() { ["reminder_key"] = key, ["title"] = title, ["deal_id"] = dealId, ["supplier_id"] = supplierId }
                }});
            WriteActivity(connection, "reminder", 0, "Удалено автоматическое напоминание", title, dealId, supplierId);
            return 0;
        });
    }

    /// <summary>
    /// Порт Set-PurchaseSupplierActualReceiptDate (PurchaseStore.ps1,
    /// строки 1894-1913): фактическая дата поступления + обновление сделки.
    /// </summary>
    public static void SetSupplierActualReceiptDate(long supplierId, DateTime? date)
    {
        if (supplierId <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(supplierId), "Выберите поставщика.");
        }
        var actualText = date?.ToString("dd.MM.yyyy") ?? string.Empty;

        PurchaseWriteSession.Execute(connection =>
        {
            var now = NowText();
            using var command = connection.CreateCommand();
            command.CommandText = "UPDATE deal_suppliers SET actual_receipt_date = @actual, updated_at = @updated_at WHERE id = @id;";
            Add(command, "@actual", actualText);
            Add(command, "@updated_at", now);
            Add(command, "@id", supplierId);
            command.ExecuteNonQuery();

            // Аналог Touch-PurchaseDealBySupplier (строки 1636-1647).
            using var touch = connection.CreateCommand();
            touch.CommandText = @"
UPDATE deals SET updated_at = @updated_at
WHERE id = (SELECT deal_id FROM deal_suppliers WHERE id = @supplier_id);";
            Add(touch, "@updated_at", now);
            Add(touch, "@supplier_id", supplierId);
            touch.ExecuteNonQuery();

            var dealId = 0L;
            using var select = connection.CreateCommand();
            select.CommandText = "SELECT deal_id FROM deal_suppliers WHERE id = @id;";
            Add(select, "@id", supplierId);
            dealId = ToLong(select.ExecuteScalar());

            WriteActivity(connection, "supplier", supplierId, "Поступление подтверждено на склад", "Факт: " + actualText, dealId, supplierId);
            return 0;
        });
    }

    /// <summary>
    /// Порт Update-ComponentDeal (PurchaseStore.ps1, строки 2372-2424) для диалога
    /// завершения задачи: статус и этап меняются, даты сработавшего вида очищаются,
    /// остальные поля сохраняются. Папка задачи не меняется, т.к. номер не редактируется.
    /// </summary>
    public static void UpdateComponentDealCompletion(long componentId, string status, string stage, bool clearReminderDate, bool clearDeadlineDate)
    {
        if (componentId <= 0)
        {
            return;
        }

        PurchaseWriteSession.Execute(connection =>
        {
            string? dealNumber = null, description = null, nextAction = null;
            string? reminderDate = null, deadlineDate = null, priority = null, period = null, previousStatus = null;
            using (var select = connection.CreateCommand())
            {
                select.CommandText = @"
SELECT IFNULL(deal_number, '') AS deal_number, IFNULL(description, '') AS description,
       IFNULL(next_action, '') AS next_action, IFNULL(reminder_date, '') AS reminder_date,
       IFNULL(deadline_date, '') AS deadline_date, IFNULL(priority, '') AS priority,
       IFNULL(period, '') AS period, IFNULL(status, '') AS status
FROM component_deals WHERE id = @id;";
                Add(select, "@id", componentId);
                using var reader = select.ExecuteReader();
                if (!reader.Read())
                {
                    throw new InvalidOperationException("Задача не найдена.");
                }
                dealNumber = reader.GetString(0);
                description = reader.GetString(1);
                nextAction = reader.GetString(2);
                reminderDate = reader.GetString(3);
                deadlineDate = reader.GetString(4);
                priority = reader.GetString(5);
                period = reader.GetString(6);
                previousStatus = reader.GetString(7);
            }

            var orderedAt = previousStatus!.Trim() != OrderedStatus && status.Trim() == OrderedStatus ? NowText() : null;
            if (clearReminderDate)
            {
                reminderDate = string.Empty;
            }
            if (clearDeadlineDate)
            {
                deadlineDate = string.Empty;
            }

            using var command = connection.CreateCommand();
            command.CommandText = @"
UPDATE component_deals SET
    deal_number = @deal_number,
    status = @status,
    ordered_at = CASE
        WHEN IFNULL(ordered_at, '') <> '' THEN ordered_at
        WHEN @ordered_at IS NOT NULL THEN @ordered_at
        ELSE ordered_at
    END,
    stage = @stage,
    description = @description,
    next_action = @next_action,
    reminder_date = @reminder_date,
    deadline_date = @deadline_date,
    priority = @priority,
    period = @period,
    updated_at = @updated_at
WHERE id = @id;";
            Add(command, "@deal_number", dealNumber);
            Add(command, "@status", status);
            Add(command, "@ordered_at", orderedAt);
            Add(command, "@stage", stage);
            Add(command, "@description", description);
            Add(command, "@next_action", nextAction);
            Add(command, "@reminder_date", reminderDate);
            Add(command, "@deadline_date", deadlineDate);
            Add(command, "@priority", priority);
            Add(command, "@period", period);
            Add(command, "@updated_at", NowText());
            Add(command, "@id", componentId);
            command.ExecuteNonQuery();

            WriteActivity(connection, "component", componentId, "Изменена задача", dealNumber!, 0);
            return 0;
        });
    }

    /// <summary>Полное сохранение редактируемой строки задачи и её заметок.</summary>
    public static void SaveComponentDeal(TaskRow task)
    {
        ArgumentNullException.ThrowIfNull(task);
        if (task.Id <= 0) throw new InvalidOperationException("Выберите задачу.");
        PurchaseWriteSession.Execute(connection =>
        {
            var previousStatus = GetString(connection, "SELECT IFNULL(status, '') FROM component_deals WHERE id = @id", task.Id);
            if (previousStatus is null) throw new InvalidOperationException("Задача не найдена.");
            var orderedAt = previousStatus.Trim() != OrderedStatus && task.Status.Trim() == OrderedStatus ? NowText() : null;
            using var command = connection.CreateCommand();
            command.CommandText = @"
UPDATE component_deals SET entry_date=@entry_date, deal_number=@deal_number, status=@status,
 ordered_at=CASE WHEN IFNULL(ordered_at,'')<>'' THEN ordered_at WHEN @ordered_at IS NOT NULL THEN @ordered_at ELSE ordered_at END,
 stage=@stage, description=@description, next_action=@next_action, reminder_date=@reminder_date,
 deadline_date=@deadline_date, priority=@priority, period=@period, order_amount=@order_amount, notes=@notes, updated_at=@updated_at WHERE id=@id;";
            Add(command, "@entry_date", task.EntryDate.Trim()); Add(command, "@deal_number", task.DealNumber.Trim());
            Add(command, "@status", task.Status.Trim()); Add(command, "@ordered_at", orderedAt); Add(command, "@stage", task.Stage.Trim());
            Add(command, "@description", task.Description.Trim()); Add(command, "@next_action", task.NextAction.Trim());
            Add(command, "@reminder_date", task.ReminderDate.Trim()); Add(command, "@deadline_date", task.DeadlineDate.Trim());
            Add(command, "@priority", task.Priority.Trim()); Add(command, "@period", task.Period.Trim());
            Add(command, "@order_amount", task.OrderAmount.Trim());
            Add(command, "@notes", task.Notes.Trim()); Add(command, "@updated_at", NowText()); Add(command, "@id", task.Id);
            command.ExecuteNonQuery();
            EnsureComponentDealFolder(connection, task.Id, task.DealNumber.Trim());
            WriteActivity(connection, "component", task.Id, "Изменена задача", task.DealNumber, 0);
            return 0;
        });
    }

    /// <summary>Сохраняет одно поле, изменённое непосредственно в строке задачи.</summary>
    public static void UpdateComponentDealGridField(long taskId, string field, string? value)
    {
        if (taskId <= 0) throw new ArgumentOutOfRangeException(nameof(taskId), "Выберите задачу.");
        var column = field switch
        {
            nameof(TaskRow.DealNumber) => "deal_number",
            nameof(TaskRow.Stage) => "stage",
            nameof(TaskRow.Description) => "description",
            nameof(TaskRow.NextAction) => "next_action",
            nameof(TaskRow.Priority) => "priority",
            nameof(TaskRow.Period) => "period",
            nameof(TaskRow.OrderAmount) => "order_amount",
            _ => null,
        };

        PurchaseWriteSession.Execute(connection =>
        {
            var now = NowText();
            using var command = connection.CreateCommand();
            if (field == nameof(TaskRow.Status))
            {
                var previousStatus = GetString(connection, "SELECT IFNULL(status, '') FROM component_deals WHERE id = @id", taskId);
                if (previousStatus is null) throw new InvalidOperationException("Задача не найдена.");
                var status = string.IsNullOrWhiteSpace(value) ? "В работе" : value.Trim();
                command.CommandText = @"
UPDATE component_deals SET status = @value,
    ordered_at = CASE WHEN IFNULL(ordered_at, '') <> '' THEN ordered_at
                      WHEN @became_ordered = 1 THEN @updated_at ELSE ordered_at END,
    updated_at = @updated_at WHERE id = @id;";
                Add(command, "@value", status);
                Add(command, "@became_ordered", previousStatus.Trim() != OrderedStatus && status == OrderedStatus ? 1 : 0);
            }
            else if (column is not null)
            {
                command.CommandText = "UPDATE component_deals SET " + column + " = @value, updated_at = @updated_at WHERE id = @id;";
                Add(command, "@value", value?.Trim() ?? string.Empty);
            }
            else
            {
                throw new ArgumentException("Поле нельзя изменить из таблицы.", nameof(field));
            }

            Add(command, "@updated_at", now);
            Add(command, "@id", taskId);
            if (command.ExecuteNonQuery() == 0) throw new InvalidOperationException("Задача не найдена.");
            if (field == nameof(TaskRow.DealNumber)) EnsureComponentDealFolder(connection, taskId, value?.Trim() ?? string.Empty);
            WriteActivity(connection, "component", taskId, "Изменено поле задачи", field, 0);
            return 0;
        });
    }

    /// <summary>
    /// Порт New-ComponentDeal (PurchaseStore.ps1, строки 2283-2298): задача с
    /// дефолтными статусом/этапом/приоритетом и собственной папкой.
    /// </summary>
    public static long CreateComponentDeal()
    {
        return PurchaseWriteSession.Execute(connection =>
        {
            var now = NowText();
            using var insert = connection.CreateCommand();
            insert.CommandText = @"
INSERT INTO component_deals(entry_date, status, stage, priority, created_at, updated_at)
VALUES(@entry_date, 'В работе', 'Запросил поставщиков', '3', @created_at, @updated_at);";
            Add(insert, "@entry_date", DateTime.Now.ToString("dd.MM.yyyy"));
            Add(insert, "@created_at", now);
            Add(insert, "@updated_at", now);
            insert.ExecuteNonQuery();

            long id;
            using (var select = connection.CreateCommand())
            {
                select.CommandText = "SELECT id FROM component_deals ORDER BY id DESC LIMIT 1;";
                id = Convert.ToInt64(select.ExecuteScalar());
            }

            var folder = EnsureComponentDealFolder(connection, id, string.Empty);
            WriteActivity(connection, "component", id, "Создана задача", folder, 0);
            return id;
        });
    }

    /// <summary>
    /// Порт Delete-ComponentDeal (PurchaseStore.ps1, строки 2449-2463): строка
    /// задачи и её папка переезжают в корзину, затем запись удаляется.
    /// </summary>
    public static void DeleteComponentDeal(long componentId)
    {
        if (componentId <= 0)
        {
            throw new InvalidOperationException("Выберите задачу.");
        }

        PurchaseWriteSession.Execute(connection =>
        {
            var rows = SelectRows(connection, "SELECT * FROM component_deals WHERE id = @id",
                command => Add(command, "@id", componentId));
            if (rows.Count == 0)
            {
                throw new InvalidOperationException("Задача не найдена.");
            }

            var componentName = Convert.ToString(rows[0].GetValueOrDefault("deal_number")) ?? string.Empty;
            var folderPath = Convert.ToString(rows[0].GetValueOrDefault("folder_path")) ?? string.Empty;
            var resolvedFolder = ResolveComponentStoredPath(folderPath);

            InsertTrashItem(connection, "component", componentId, "Задача: " + componentName,
                new Dictionary<string, object> { ["component_deals"] = rows },
                null,
                resolvedFolder);

            using var delete = connection.CreateCommand();
            delete.CommandText = "DELETE FROM component_deals WHERE id = @id;";
            Add(delete, "@id", componentId);
            delete.ExecuteNonQuery();

            RemoveComponentDealFolder(folderPath);
            WriteActivity(connection, "component", componentId, "Удалена задача", string.Empty, 0);
            return 0;
        });
    }

    /// <summary>Корень папок задач: {данные}\components\files (Get-DefaultComponentFilesRoot).</summary>
    private static string ComponentFilesRoot => Path.Combine(AppPaths.DataRoot, "components", "files");

    /// <summary>Порт Get-ComponentDealFolderPath (PurchaseStore.ps1, строки 1422-1432).</summary>
    private static string GetComponentDealFolderPath(long id, string dealNumber)
    {
        var name = string.IsNullOrWhiteSpace(dealNumber)
            ? $"component_{id:D5}"
            : $"{id:D5}_{PurchaseDocumentsStore.SafePathPart(dealNumber)}";
        return Path.Combine(ComponentFilesRoot, name);
    }

    /// <summary>Порт Resolve-ComponentStoredPath (строки 119-133) для папки.</summary>
    private static string ResolveComponentStoredPath(string? storedPath)
    {
        if (string.IsNullOrWhiteSpace(storedPath))
        {
            return string.Empty;
        }
        var path = ResolvePortable(storedPath);
        if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
        {
            try
            {
                return Path.GetFullPath(path);
            }
            catch (ArgumentException)
            {
                return path;
            }
        }

        var candidate = PurchaseDocumentsStore.GetMovedTreeCandidate(storedPath, Path.Combine("data", "components", "files"));
        if (!string.IsNullOrWhiteSpace(candidate) && Directory.Exists(candidate))
        {
            try
            {
                return Path.GetFullPath(candidate);
            }
            catch (ArgumentException)
            {
                return candidate;
            }
        }
        return path;
    }

    /// <summary>
    /// Порт Ensure-ComponentDealFolder (строки 1434-1473): при смене номера
    /// сделки папка переезжает, folder_path хранится портативно.
    /// </summary>
    private static string EnsureComponentDealFolder(SqliteConnection connection, long id, string dealNumber)
    {
        if (id <= 0)
        {
            return string.Empty;
        }

        Directory.CreateDirectory(ComponentFilesRoot);
        var target = GetComponentDealFolderPath(id, dealNumber);
        var currentStored = GetString(connection, "SELECT IFNULL(folder_path, '') FROM component_deals WHERE id = @id", id) ?? string.Empty;
        var current = ResolveComponentStoredPath(currentStored);
        var chosen = target;
        var samePath = false;
        if (!string.IsNullOrWhiteSpace(current))
        {
            try
            {
                samePath = Path.GetFullPath(current).TrimEnd('\\') == Path.GetFullPath(target).TrimEnd('\\');
            }
            catch (ArgumentException)
            {
                samePath = current == target;
            }
        }

        if (!string.IsNullOrWhiteSpace(current) && !samePath)
        {
            try
            {
                if (Directory.Exists(current) && !Directory.Exists(target))
                {
                    Directory.Move(current, target);
                    chosen = target;
                }
                else if (Directory.Exists(current))
                {
                    chosen = current;
                }
            }
            catch (IOException)
            {
                chosen = current;
            }
        }

        if (string.IsNullOrWhiteSpace(chosen))
        {
            chosen = target;
        }
        Directory.CreateDirectory(chosen);

        using var update = connection.CreateCommand();
        update.CommandText = "UPDATE component_deals SET folder_path = @folder_path WHERE id = @id;";
        Add(update, "@folder_path", PurchaseDocumentsStore.ToPortablePath(chosen));
        Add(update, "@id", id);
        update.ExecuteNonQuery();
        return chosen;
    }

    /// <summary>
    /// Порт Remove-ComponentDealFolder (строки 1475-1491): удаляет папку задачи
    /// только если она лежит внутри корня папок задач.
    /// </summary>
    private static void RemoveComponentDealFolder(string folderPath)
    {
        if (string.IsNullOrWhiteSpace(folderPath))
        {
            return;
        }

        string root;
        string target;
        try
        {
            root = Path.GetFullPath(ComponentFilesRoot);
            target = Path.GetFullPath(ResolveComponentStoredPath(folderPath));
        }
        catch (ArgumentException)
        {
            return;
        }
        if (!target.StartsWith(root, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }
        if (Directory.Exists(target))
        {
            Directory.Delete(target, recursive: true);
        }
    }

    /// <summary>
    /// Порт Set-NotificationState (PurchaseStore.ps1, строки 1031-1063).
    /// </summary>
    public static void SetNotificationState(
        string source,
        long sourceId,
        string dueKind,
        string dueDate,
        bool handled,
        string? snoozeUntil = null,
        bool shown = false)
    {
        PurchaseWriteSession.Execute(connection =>
        {
            var now = NowText();
            using var command = connection.CreateCommand();
            command.CommandText = @"
INSERT INTO notification_state(source, source_id, due_kind, due_date, handled, snooze_until, last_shown_at, created_at, updated_at)
VALUES(@source, @source_id, @due_kind, @due_date, @handled, @snooze_until, @last_shown_at, @created_at, @updated_at)
ON CONFLICT(source, source_id, due_kind, due_date) DO UPDATE SET
    handled = @handled,
    snooze_until = @snooze_until,
    last_shown_at = CASE WHEN @last_shown_at IS NULL THEN notification_state.last_shown_at ELSE @last_shown_at END,
    updated_at = @updated_at;";
            Add(command, "@source", source);
            Add(command, "@source_id", sourceId);
            Add(command, "@due_kind", dueKind);
            Add(command, "@due_date", dueDate);
            Add(command, "@handled", handled ? 1 : 0);
            Add(command, "@snooze_until", string.IsNullOrWhiteSpace(snoozeUntil) ? null : snoozeUntil);
            Add(command, "@last_shown_at", shown ? now : null);
            Add(command, "@created_at", now);
            Add(command, "@updated_at", now);
            command.ExecuteNonQuery();
            return 0;
        });
    }

    /// <summary>
    /// Порт Remove-QuoteHistoryItems (QuoteHistory.ps1, строки 199-227): квоты и их
    /// батчи сохраняются в корзину, затем удаляются; пустые батчи подчищаются.
    /// </summary>
    public static int RemoveQuoteHistoryItems(IReadOnlyList<long> ids)
    {
        var validIds = ids.Where(id => id > 0).Distinct().ToList();
        if (validIds.Count == 0)
        {
            return 0;
        }

        PurchaseWriteSession.Execute(connection =>
        {
            var quoteRows = new List<Dictionary<string, object?>>();
            foreach (var id in validIds)
            {
                quoteRows.AddRange(SelectRows(connection, "SELECT * FROM quote_history WHERE id = @id",
                    command => Add(command, "@id", id)));
            }

            var batchRows = new List<Dictionary<string, object?>>();
            var batchIds = quoteRows
                .Select(row => row.GetValueOrDefault("batch_id"))
                .Where(value => value is not null)
                .Select(Convert.ToInt64)
                .Distinct();
            foreach (var batchId in batchIds)
            {
                batchRows.AddRange(SelectRows(connection, "SELECT * FROM quote_batches WHERE id = @id",
                    command => Add(command, "@id", batchId)));
            }

            InsertTrashItem(connection, "quote_history", 0, "Квоты: " + validIds.Count, new Dictionary<string, object>
            {
                ["quote_batches"] = batchRows,
                ["quote_history"] = quoteRows,
            });

            foreach (var id in validIds)
            {
                using var delete = connection.CreateCommand();
                delete.CommandText = "DELETE FROM quote_history WHERE id = @id;";
                Add(delete, "@id", id);
                delete.ExecuteNonQuery();
            }

            using var cleanBatches = connection.CreateCommand();
            cleanBatches.CommandText = "DELETE FROM quote_batches WHERE id NOT IN (SELECT DISTINCT batch_id FROM quote_history);";
            cleanBatches.ExecuteNonQuery();

            WriteActivity(connection, "quote_history", 0, "Удалены квоты из базы квот", "Количество: " + validIds.Count, 0);
            return 0;
        });
        return validIds.Count;
    }

    /// <summary>
    /// Порт Clear-QuoteHistory (QuoteHistory.ps1, строки 229-238): вся база квот
    /// сохраняется в корзину, затем обе таблицы очищаются.
    /// </summary>
    public static void ClearQuoteHistory()
    {
        PurchaseWriteSession.Execute(connection =>
        {
            var quoteRows = SelectRows(connection, "SELECT * FROM quote_history");
            var batchRows = SelectRows(connection, "SELECT * FROM quote_batches");
            InsertTrashItem(connection, "quote_history", 0, "Очищена база квот", new Dictionary<string, object>
            {
                ["quote_batches"] = batchRows,
                ["quote_history"] = quoteRows,
            });

            using var deleteQuotes = connection.CreateCommand();
            deleteQuotes.CommandText = "DELETE FROM quote_history;";
            deleteQuotes.ExecuteNonQuery();

            using var deleteBatches = connection.CreateCommand();
            deleteBatches.CommandText = "DELETE FROM quote_batches;";
            deleteBatches.ExecuteNonQuery();
            return 0;
        });
    }

    /// <summary>Совместимый импорт одиночных документов; полноценные когорты сохраняются через SaveAnalysisSnapshot.</summary>
    public static int SaveQuoteHistory(IReadOnlyList<Quote> quotes, string rfqPath, string comment)
    {
        ArgumentNullException.ThrowIfNull(quotes);
        if (quotes.Count == 0) return 0;

        var fingerprint = "document:" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(
            rfqPath + "\n" + string.Join("\n", quotes.OrderBy(quote => quote.WorkbookPath, StringComparer.OrdinalIgnoreCase)
                .ThenBy(quote => quote.SheetName, StringComparer.OrdinalIgnoreCase).ThenBy(quote => quote.Row)
                .Select(quote => $"{quote.WorkbookPath}|{quote.SheetName}|{quote.Row}|{quote.Supplier}|{quote.Key}|{quote.UnitPrice}|{quote.LeadTime}")))));

        return PurchaseWriteSession.Execute(connection =>
        {
            using (var find = connection.CreateCommand())
            {
                find.CommandText = "SELECT id FROM quote_batches WHERE fingerprint=@fingerprint LIMIT 1;";
                Add(find, "@fingerprint", fingerprint);
                if (find.ExecuteScalar() is not null and not DBNull) return 0;
            }

            var now = NowText();
            using (var batch = connection.CreateCommand())
            {
                batch.CommandText = @"
INSERT INTO quote_batches(created_at, rfq_path, priority, comment, source_kind, fingerprint, data_quality)
VALUES(@created_at, @rfq_path, 'Price', @comment, 'document_import', @fingerprint, 'incomplete');";
                Add(batch, "@created_at", now); Add(batch, "@rfq_path", rfqPath); Add(batch, "@comment", comment); Add(batch, "@fingerprint", fingerprint);
                batch.ExecuteNonQuery();
            }
            long batchId;
            using (var id = connection.CreateCommand())
            {
                id.CommandText = "SELECT last_insert_rowid();";
                batchId = Convert.ToInt64(id.ExecuteScalar());
            }

            var saved = 0;
            foreach (var quote in quotes)
            {
                using var duplicate = connection.CreateCommand();
                duplicate.CommandText = @"
SELECT COUNT(*) FROM quote_history
WHERE batch_id = @batch_id
  AND lower(trim(IFNULL(supplier, ''))) = lower(trim(@supplier))
  AND IFNULL(rfq_value, '') = IFNULL(@rfq_value, '')
  AND IFNULL(pn, '') = IFNULL(@pn, '')
  AND IFNULL(unit_price, -999999999) = IFNULL(@unit_price, -999999999)
  AND IFNULL(lead_time, '') = IFNULL(@lead_time, '');";
                Add(duplicate, "@batch_id", batchId); Add(duplicate, "@supplier", quote.Supplier); Add(duplicate, "@rfq_value", quote.Key); Add(duplicate, "@pn", quote.PN);
                Add(duplicate, "@unit_price", quote.UnitPrice); Add(duplicate, "@lead_time", quote.LeadTime);
                if (Convert.ToInt64(duplicate.ExecuteScalar()) > 0) continue;

                using var insert = connection.CreateCommand();
                insert.CommandText = @"
INSERT INTO quote_history(batch_id, quote_date, rfq_value, pn, supplier, unit_price, lead_time, lead_time_total, mfg, is_winner, warning, sheet_name, row_number, match_status, winner_reason, requested_qty)
VALUES(@batch_id, @quote_date, @rfq_value, @pn, @supplier, @unit_price, @lead_time, @lead_time_total, @mfg, 0, @warning, @sheet_name, @row_number, @match_status, '', @requested_qty);";
                Add(insert, "@batch_id", batchId); Add(insert, "@quote_date", now); Add(insert, "@rfq_value", quote.Key);
                Add(insert, "@pn", quote.PN); Add(insert, "@supplier", quote.Supplier); Add(insert, "@unit_price", quote.UnitPrice);
                Add(insert, "@lead_time", quote.LeadTime); Add(insert, "@lead_time_total", quote.LeadTimeTotal);
                var legacyManufacturer = string.IsNullOrWhiteSpace(quote.MfgChina) ? quote.MfgRussia : quote.MfgChina;
                Add(insert, "@mfg", legacyManufacturer);
                AddManufacturerAliasCandidate(connection, legacyManufacturer);
                Add(insert, "@warning", quote.Warning); Add(insert, "@sheet_name", quote.SheetName);
                Add(insert, "@row_number", quote.Row == 0 ? null : quote.Row); Add(insert, "@match_status", quote.MatchStatus);
                Add(insert, "@requested_qty", quote.QtyToBuy);
                insert.ExecuteNonQuery(); saved++;
            }

            if (saved == 0)
            {
                using var removeBatch = connection.CreateCommand();
                removeBatch.CommandText = "DELETE FROM quote_batches WHERE id = @id;";
                Add(removeBatch, "@id", batchId); removeBatch.ExecuteNonQuery();
            }
            else
            {
                WriteActivity(connection, "rrfq", batchId, "Сохранена база квот", "Квот: " + saved, 0);
            }
            return saved;
        });
    }

    /// <summary>
    /// Сохраняет полный снимок одного запуска сравнения RRFQ. В отличие от
    /// старого SaveQuoteHistory, здесь сохраняются позиции, участники, все
    /// предложения и фактический победитель.
    /// </summary>
    public static AnalysisSaveResult SaveAnalysisSnapshot(RrfqAnalysis analysis, string resultPath)
    {
        ArgumentNullException.ThrowIfNull(analysis);
        var fingerprint = BuildAnalysisFingerprint(analysis);

        return PurchaseWriteSession.Execute(connection =>
        {
            var now = NowText();
            long batchId;
            var duplicate = false;
            using (var find = connection.CreateCommand())
            {
                find.CommandText = "SELECT id FROM quote_batches WHERE fingerprint = @fingerprint ORDER BY id DESC LIMIT 1;";
                Add(find, "@fingerprint", fingerprint);
                var value = find.ExecuteScalar();
                if (value is not null and not DBNull)
                {
                    batchId = Convert.ToInt64(value);
                    duplicate = true;
                    using var update = connection.CreateCommand();
                    update.CommandText = @"UPDATE quote_batches SET rfq_path=@rfq_path, result_path=@result_path,
priority=@priority, source_kind='full_analysis', data_quality='full', comment=@comment WHERE id=@id;";
                    Add(update, "@rfq_path", analysis.RfqPath);
                    Add(update, "@result_path", resultPath);
                    Add(update, "@priority", analysis.Priority);
                    Add(update, "@comment", "analysis");
                    Add(update, "@id", batchId);
                    update.ExecuteNonQuery();

                    DeleteBatchSnapshotRows(connection, batchId);
                }
                else
                {
                    using var insert = connection.CreateCommand();
                    insert.CommandText = @"INSERT INTO quote_batches(created_at, rfq_path, priority, comment, result_path, source_kind, fingerprint, data_quality)
VALUES(@created_at, @rfq_path, @priority, 'analysis', @result_path, 'full_analysis', @fingerprint, 'full');";
                    Add(insert, "@created_at", now);
                    Add(insert, "@rfq_path", analysis.RfqPath);
                    Add(insert, "@priority", analysis.Priority);
                    Add(insert, "@result_path", resultPath);
                    Add(insert, "@fingerprint", fingerprint);
                    insert.ExecuteNonQuery();
                    using var id = connection.CreateCommand();
                    id.CommandText = "SELECT last_insert_rowid();";
                    batchId = Convert.ToInt64(id.ExecuteScalar());
                }
            }

            var rowsById = analysis.RfqRows.ToDictionary(row => row.Id, StringComparer.Ordinal);
            var quoteManufacturerByTarget = analysis.Quotes
                .Where(quote => !string.IsNullOrWhiteSpace(quote.TargetId))
                .GroupBy(quote => quote.TargetId!, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => FirstNonEmpty(group.Select(quote => FirstNonEmpty(quote.MfgChina, quote.MfgRussia)).ToArray()), StringComparer.Ordinal);
            foreach (var row in analysis.RfqRows)
            {
                var manufacturer = FirstNonEmpty(row.MfgChina, row.MfgRussia,
                    quoteManufacturerByTarget.GetValueOrDefault(row.Id, string.Empty));
                AddManufacturerAliasCandidate(connection, manufacturer);
                InsertBatchPosition(connection, batchId, row.Id, row.SheetName, row.Row, row.Key,
                    row.PN, manufacturer, row.QtyToBuy, row.RussianRemark);
            }

            foreach (var supplier in analysis.Suppliers.Where(item => !string.IsNullOrWhiteSpace(item.Supplier)))
            {
                using var insert = connection.CreateCommand();
                insert.CommandText = @"INSERT OR IGNORE INTO quote_batch_suppliers(batch_id, supplier, source_path)
VALUES(@batch_id, @supplier, @source_path);";
                Add(insert, "@batch_id", batchId);
                Add(insert, "@supplier", supplier.Supplier.Trim());
                Add(insert, "@source_path", supplier.Path);
                insert.ExecuteNonQuery();
            }

            var winners = analysis.Decisions
                .Where(decision => decision.Include && decision.Winner is not null)
                .ToDictionary(decision => decision.Winner!.Id,
                    decision => (decision.WinnerReason, decision.Id), StringComparer.Ordinal);
            var saved = 0;
            foreach (var quote in analysis.Quotes)
            {
                rowsById.TryGetValue(quote.TargetId ?? string.Empty, out var target);
                var positionKey = quote.TargetId;
                if (string.IsNullOrWhiteSpace(positionKey))
                    positionKey = $"unmatched:{quote.SheetName}:{quote.Row}:{quote.KeyNorm}";

                var manufacturer = FirstNonEmpty(quote.MfgChina, quote.MfgRussia,
                    target is null ? string.Empty : FirstNonEmpty(target.MfgChina, target.MfgRussia));
                AddManufacturerAliasCandidate(connection, manufacturer);
                winners.TryGetValue(quote.Id, out var winner);

                using var insert = connection.CreateCommand();
                insert.CommandText = @"INSERT INTO quote_history(
batch_id, quote_date, rfq_value, pn, supplier, unit_price, lead_time, lead_time_total, mfg,
is_winner, winner_reason, warning, sheet_name, row_number, match_status, position_key,
quote_id, currency, requested_qty, manufacturer_raw, manufacturer_norm, lead_total_days, is_price_comparable)
VALUES(@batch_id, @quote_date, @rfq_value, @pn, @supplier, @unit_price, @lead_time, @lead_time_total, @mfg,
@is_winner, @winner_reason, @warning, @sheet_name, @row_number, @match_status, @position_key,
@quote_id, @currency, @requested_qty, @manufacturer_raw, @manufacturer_norm, @lead_total_days, @is_price_comparable);";
                Add(insert, "@batch_id", batchId);
                Add(insert, "@quote_date", now);
                Add(insert, "@rfq_value", quote.Key);
                Add(insert, "@pn", quote.PN);
                Add(insert, "@supplier", quote.Supplier);
                Add(insert, "@unit_price", quote.UnitPrice);
                Add(insert, "@lead_time", quote.LeadTime);
                Add(insert, "@lead_time_total", quote.LeadTimeTotal);
                Add(insert, "@mfg", manufacturer);
                Add(insert, "@is_winner", winner != default ? 1 : 0);
                Add(insert, "@winner_reason", winner != default ? winner.WinnerReason : string.Empty);
                Add(insert, "@warning", quote.Warning);
                Add(insert, "@sheet_name", quote.SheetName);
                Add(insert, "@row_number", quote.Row == 0 ? null : quote.Row);
                Add(insert, "@match_status", quote.MatchStatus);
                Add(insert, "@position_key", positionKey);
                Add(insert, "@quote_id", quote.Id);
                Add(insert, "@currency", quote.Currency);
                Add(insert, "@requested_qty", target?.QtyToBuy ?? quote.QtyToBuy);
                Add(insert, "@manufacturer_raw", manufacturer);
                Add(insert, "@manufacturer_norm", SupplierAnalyticsService.NormalizeManufacturer(manufacturer));
                Add(insert, "@lead_total_days", SupplierAnalyticsService.ParseLeadTimeTotalDays(quote.LeadTimeTotal));
                Add(insert, "@is_price_comparable", quote.IsPriceComparable ? 1 : 0);
                insert.ExecuteNonQuery();
                saved++;
            }

            WriteActivity(connection, "rrfq", batchId,
                duplicate ? "Обновлена аналитическая когорта RRFQ" : "Сохранена аналитическая когорта RRFQ",
                "Квот: " + saved, 0);
            return new AnalysisSaveResult(batchId, saved, duplicate);
        });
    }

    /// <summary>Импорт старого результирующего RRFQ, где сохранены только победители.</summary>
    public static AnalysisSaveResult SaveWinnerOnlyQuotes(IReadOnlyList<Quote> quotes, string path)
    {
        ArgumentNullException.ThrowIfNull(quotes);
        if (quotes.Count == 0) return new AnalysisSaveResult(0, 0, false, true);
        var fingerprint = "winner-only:" + FileFingerprint(path);

        return PurchaseWriteSession.Execute(connection =>
        {
            using var find = connection.CreateCommand();
            find.CommandText = "SELECT id FROM quote_batches WHERE fingerprint=@fingerprint LIMIT 1;";
            Add(find, "@fingerprint", fingerprint);
            var existing = find.ExecuteScalar();
            if (existing is not null and not DBNull)
                return new AnalysisSaveResult(Convert.ToInt64(existing), 0, true, true);

            using var batch = connection.CreateCommand();
            batch.CommandText = @"INSERT INTO quote_batches(created_at, rfq_path, priority, comment, result_path, source_kind, fingerprint, data_quality)
VALUES(@created_at, @rfq_path, 'Price', 'winner-only import', @result_path, 'winner-only', @fingerprint, 'incomplete');";
            Add(batch, "@created_at", NowText()); Add(batch, "@rfq_path", path); Add(batch, "@result_path", path); Add(batch, "@fingerprint", fingerprint);
            batch.ExecuteNonQuery();
            using var idCommand = connection.CreateCommand();
            idCommand.CommandText = "SELECT last_insert_rowid();";
            var batchId = Convert.ToInt64(idCommand.ExecuteScalar());

            foreach (var group in quotes.GroupBy(quote => quote.SheetName + ":" + quote.Row + ":" + quote.KeyNorm))
            {
                var first = group.First();
                var positionKey = "winner-only:" + group.Key;
                var manufacturer = FirstNonEmpty(first.MfgChina, first.MfgRussia);
                AddManufacturerAliasCandidate(connection, manufacturer);
                InsertBatchPosition(connection, batchId, positionKey, first.SheetName, first.Row, first.Key, first.PN, manufacturer, first.QtyToBuy, first.RussianRemark);
                using var supplier = connection.CreateCommand();
                supplier.CommandText = "INSERT OR IGNORE INTO quote_batch_suppliers(batch_id, supplier, source_path) VALUES(@batch_id,@supplier,@source_path);";
                Add(supplier, "@batch_id", batchId); Add(supplier, "@supplier", first.Supplier); Add(supplier, "@source_path", path);
                supplier.ExecuteNonQuery();

                foreach (var quote in group)
                {
                    using var insert = connection.CreateCommand();
                    insert.CommandText = @"INSERT INTO quote_history(
batch_id, quote_date, rfq_value, pn, supplier, unit_price, lead_time, lead_time_total, mfg,
is_winner, winner_reason, warning, sheet_name, row_number, match_status, position_key, quote_id,
currency, requested_qty, manufacturer_raw, manufacturer_norm, lead_total_days, is_price_comparable)
VALUES(@batch_id,@quote_date,@rfq_value,@pn,@supplier,@unit_price,@lead_time,@lead_time_total,@mfg,
1,'winner-only',@warning,@sheet_name,@row_number,@match_status,@position_key,@quote_id,
@currency,@requested_qty,@manufacturer_raw,@manufacturer_norm,@lead_total_days,@comparable);";
                    Add(insert, "@batch_id", batchId); Add(insert, "@quote_date", NowText()); Add(insert, "@rfq_value", quote.Key);
                    Add(insert, "@pn", quote.PN); Add(insert, "@supplier", quote.Supplier); Add(insert, "@unit_price", quote.UnitPrice);
                    Add(insert, "@lead_time", quote.LeadTime); Add(insert, "@lead_time_total", quote.LeadTimeTotal); Add(insert, "@mfg", manufacturer);
                    Add(insert, "@warning", quote.Warning); Add(insert, "@sheet_name", quote.SheetName); Add(insert, "@row_number", quote.Row);
                    Add(insert, "@match_status", "winner-only"); Add(insert, "@position_key", positionKey); Add(insert, "@quote_id", quote.Id);
                    Add(insert, "@currency", quote.Currency); Add(insert, "@requested_qty", quote.QtyToBuy); Add(insert, "@manufacturer_raw", manufacturer);
                    Add(insert, "@manufacturer_norm", SupplierAnalyticsService.NormalizeManufacturer(manufacturer));
                    Add(insert, "@lead_total_days", SupplierAnalyticsService.ParseLeadTimeTotalDays(quote.LeadTimeTotal));
                    Add(insert, "@comparable", quote.IsPriceComparable ? 1 : 0); insert.ExecuteNonQuery();
                }
            }

            WriteActivity(connection, "rrfq", batchId, "Импортирован результирующий RRFQ", "Квот: " + quotes.Count, 0);
            return new AnalysisSaveResult(batchId, quotes.Count, false, true);
        });
    }

    private static void DeleteBatchSnapshotRows(SqliteConnection connection, long batchId)
    {
        foreach (var table in new[] { "quote_history", "quote_batch_positions", "quote_batch_suppliers" })
        {
            using var delete = connection.CreateCommand();
            delete.CommandText = $"DELETE FROM {table} WHERE batch_id = @batch_id;";
            Add(delete, "@batch_id", batchId);
            delete.ExecuteNonQuery();
        }
    }

    private static void InsertBatchPosition(SqliteConnection connection, long batchId, string positionKey,
        string sheetName, int rowNumber, string rfqValue, string pn, string manufacturer, string quantity, string description)
    {
        using var insert = connection.CreateCommand();
        insert.CommandText = @"INSERT OR REPLACE INTO quote_batch_positions(
batch_id, position_key, sheet_name, row_number, rfq_value, pn, manufacturer_raw, manufacturer_norm, requested_qty, description)
VALUES(@batch_id, @position_key, @sheet_name, @row_number, @rfq_value, @pn, @manufacturer_raw, @manufacturer_norm, @requested_qty, @description);";
        Add(insert, "@batch_id", batchId);
        Add(insert, "@position_key", positionKey);
        Add(insert, "@sheet_name", sheetName);
        Add(insert, "@row_number", rowNumber);
        Add(insert, "@rfq_value", rfqValue);
        Add(insert, "@pn", pn);
        Add(insert, "@manufacturer_raw", manufacturer);
        Add(insert, "@manufacturer_norm", SupplierAnalyticsService.NormalizeManufacturer(manufacturer));
        Add(insert, "@requested_qty", quantity);
        Add(insert, "@description", description);
        insert.ExecuteNonQuery();
    }

    private static void AddManufacturerAliasCandidate(SqliteConnection connection, string manufacturer)
    {
        var raw = manufacturer.Trim();
        var normalized = SupplierAnalyticsService.NormalizeManufacturer(raw);
        if (normalized.Length == 0) return;

        using var insert = connection.CreateCommand();
        insert.CommandText = @"INSERT OR IGNORE INTO manufacturer_aliases(manufacturer_id, alias, normalized_alias, created_at, updated_at)
VALUES((SELECT id FROM manufacturers WHERE normalized_name=@normalized), @alias, @normalized, @created_at, @updated_at);";
        Add(insert, "@normalized", normalized);
        Add(insert, "@alias", raw);
        Add(insert, "@created_at", NowText());
        Add(insert, "@updated_at", NowText());
        insert.ExecuteNonQuery();
    }

    private static string BuildAnalysisFingerprint(RrfqAnalysis analysis)
    {
        using var sha = SHA256.Create();
        var parts = new List<string> { "full_analysis", analysis.Priority, FileFingerprint(analysis.RfqPath) };
        parts.AddRange(analysis.Suppliers
            .OrderBy(item => item.Supplier, StringComparer.OrdinalIgnoreCase)
            .ThenBy(item => item.Path, StringComparer.OrdinalIgnoreCase)
            .Select(item => item.Supplier.Trim() + "|" + FileFingerprint(item.Path)));
        return Convert.ToHexString(sha.ComputeHash(Encoding.UTF8.GetBytes(string.Join("\n", parts))));
    }

    private static string FileFingerprint(string path)
    {
        if (!File.Exists(path)) return "missing:" + path;
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private static string FirstNonEmpty(params string[] values)
        => values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? string.Empty;

    /// <summary>
    /// Финал порта Import-GlobalistWorkbook (QuoteHistory.ps1, строки 152-160):
    /// таблица globalist_quotes полностью заменяется разобранными строками.
    /// </summary>
    public static void ReplaceGlobalistRows(IReadOnlyList<GlobalistImportRecord> records)
    {
        ArgumentNullException.ThrowIfNull(records);

        PurchaseWriteSession.Execute(connection =>
        {
            using var delete = connection.CreateCommand();
            delete.CommandText = "DELETE FROM globalist_quotes;";
            delete.ExecuteNonQuery();

            foreach (var record in records)
            {
                using var command = connection.CreateCommand();
                command.CommandText = @"
INSERT INTO globalist_quotes(imported_at, source_file, sheet_name, row_number, factory, pn, comment, pi_number, replacement, chinese_remark, package, brand, datacode, moq, qty, stock, need_spq, spq, unit_price, total_amount, lead_time, weight, target, supplier_quote_id)
VALUES(@imported_at, @source_file, @sheet_name, @row_number, @factory, @pn, @comment, @pi_number, @replacement, @chinese_remark, @package, @brand, @datacode, @moq, @qty, @stock, @need_spq, @spq, @unit_price, @total_amount, @lead_time, @weight, @target, @supplier_quote_id);";
                Add(command, "@imported_at", record.ImportedAt);
                Add(command, "@source_file", record.SourceFile);
                Add(command, "@sheet_name", record.SheetName);
                Add(command, "@row_number", record.RowNumber);
                Add(command, "@factory", record.Factory);
                Add(command, "@pn", record.Pn);
                Add(command, "@comment", record.Comment);
                Add(command, "@pi_number", record.PiNumber);
                Add(command, "@replacement", record.Replacement);
                Add(command, "@chinese_remark", record.ChineseRemark);
                Add(command, "@package", record.Package);
                Add(command, "@brand", record.Brand);
                Add(command, "@datacode", record.Datacode);
                Add(command, "@moq", record.Moq);
                Add(command, "@qty", record.Qty);
                Add(command, "@stock", record.Stock);
                Add(command, "@need_spq", record.NeedSpq);
                Add(command, "@spq", record.Spq);
                Add(command, "@unit_price", record.UnitPrice);
                Add(command, "@total_amount", record.TotalAmount);
                Add(command, "@lead_time", record.LeadTime);
                Add(command, "@weight", record.Weight);
                Add(command, "@target", record.Target);
                Add(command, "@supplier_quote_id", record.SupplierQuoteId);
                command.ExecuteNonQuery();
            }
            return 0;
        });
    }

    private static readonly string[] TrashRestoreTables =
    {
        "deals", "component_deals", "deal_suppliers", "reminders", "reminder_suppressions", "documents",
        "quote_batches", "quote_history", "activity_log", "bitrix_task_links",
    };

    private static readonly Regex SafeColumnName = new("^[A-Za-z_][A-Za-z0-9_]*$", RegexOptions.Compiled);

    /// <summary>
    /// Порт Restore-PurchaseTrashItem (PurchaseStore.ps1, строки 779-806): строки из
    /// payload_json возвращаются через INSERT OR REPLACE, файлы переезжают обратно,
    /// папка и запись корзины удаляются.
    /// </summary>
    public static void RestoreTrashItem(long trashId)
    {
        if (trashId <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(trashId), "Выберите запись в корзине.");
        }

        PurchaseWriteSession.Execute(connection =>
        {
            string payloadJson, filesJson, title;
            using (var select = connection.CreateCommand())
            {
                select.CommandText = "SELECT payload_json, files_json, title FROM trash_items WHERE id = @id;";
                Add(select, "@id", trashId);
                using var reader = select.ExecuteReader();
                if (!reader.Read())
                {
                    throw new InvalidOperationException("Запись корзины не найдена.");
                }
                payloadJson = reader.IsDBNull(0) ? string.Empty : reader.GetString(0);
                filesJson = reader.IsDBNull(1) ? string.Empty : reader.GetString(1);
                title = reader.IsDBNull(2) ? string.Empty : reader.GetString(2);
            }

            if (!string.IsNullOrWhiteSpace(payloadJson))
            {
                using var document = JsonDocument.Parse(payloadJson);
                foreach (var tableName in TrashRestoreTables)
                {
                    if (!document.RootElement.TryGetProperty(tableName, out var rowsElement) ||
                        rowsElement.ValueKind != JsonValueKind.Array)
                    {
                        continue;
                    }

                    foreach (var rowElement in rowsElement.EnumerateArray())
                    {
                        RestoreTrashRow(connection, tableName, rowElement);
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(filesJson))
            {
                using var document = JsonDocument.Parse(filesJson);
                foreach (var record in document.RootElement.EnumerateArray())
                {
                    var source = ResolvePortable(record.GetProperty("TrashPath").GetString());
                    var target = ResolvePortable(record.GetProperty("OriginalPath").GetString());
                    if (string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(target) ||
                        !File.Exists(source) && !Directory.Exists(source))
                    {
                        continue;
                    }

                    var parent = Path.GetDirectoryName(target);
                    if (!string.IsNullOrWhiteSpace(parent))
                    {
                        Directory.CreateDirectory(parent);
                    }

                    if (Directory.Exists(source))
                    {
                        if (Directory.Exists(target))
                        {
                            Directory.Delete(target, recursive: true);
                        }
                        Directory.Move(source, target);
                    }
                    else
                    {
                        File.Move(source, target, overwrite: true);
                    }
                }
            }

            var itemRoot = Path.Combine(AppPaths.PurchaseDataDirectory, "trash_files", trashId.ToString());
            if (Directory.Exists(itemRoot))
            {
                Directory.Delete(itemRoot, recursive: true);
            }

            using var delete = connection.CreateCommand();
            delete.CommandText = "DELETE FROM trash_items WHERE id = @id;";
            Add(delete, "@id", trashId);
            delete.ExecuteNonQuery();

            WriteActivity(connection, "trash", trashId, "Восстановлено из корзины", title, 0);
            return 0;
        });
    }

    /// <summary>
    /// Порт Remove-PurchaseTrashItem (PurchaseStore.ps1, строки 808-815): безвозвратное удаление.
    /// </summary>
    public static void RemoveTrashItem(long trashId)
    {
        if (trashId <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(trashId), "Выберите запись в корзине.");
        }

        PurchaseWriteSession.Execute(connection =>
        {
            using var delete = connection.CreateCommand();
            delete.CommandText = "DELETE FROM trash_items WHERE id = @id;";
            Add(delete, "@id", trashId);
            delete.ExecuteNonQuery();
            return 0;
        });

        var itemRoot = Path.Combine(AppPaths.PurchaseDataDirectory, "trash_files", trashId.ToString());
        if (Directory.Exists(itemRoot))
        {
            Directory.Delete(itemRoot, recursive: true);
        }
    }

    public static void MoveFileToTrash(string entityType, string title, string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            throw new FileNotFoundException("Файл для перемещения в корзину не найден.", filePath);
        }

        PurchaseWriteSession.Execute(connection =>
        {
            var trashId = InsertTrashItem(
                connection,
                entityType,
                0,
                title,
                new Dictionary<string, object>(),
                new[] { filePath },
                string.Empty);
            WriteActivity(connection, "trash", trashId, "Перемещено в корзину", title, 0);
            return 0;
        });
    }

    /// <summary>Одна строка из payload: INSERT OR REPLACE с проверкой имён колонок (как в Restore-PurchaseTrashRows).</summary>
    private static void RestoreTrashRow(SqliteConnection connection, string tableName, JsonElement rowElement)
    {
        if (rowElement.ValueKind != JsonValueKind.Object)
        {
            return;
        }

        var columns = new List<string>();
        var values = new List<object?>();
        foreach (var property in rowElement.EnumerateObject())
        {
            if (!SafeColumnName.IsMatch(property.Name))
            {
                continue;
            }
            columns.Add(property.Name);
            values.Add(JsonToDb(property.Value));
        }
        if (columns.Count == 0)
        {
            return;
        }

        var names = columns.Select(column => "@p_" + column).ToList();
        using var command = connection.CreateCommand();
        command.CommandText = $"INSERT OR REPLACE INTO {tableName} ({string.Join(", ", columns)}) VALUES ({string.Join(", ", names)});";
        for (var i = 0; i < columns.Count; i++)
        {
            Add(command, names[i], values[i]);
        }
        command.ExecuteNonQuery();
    }

    private static object? JsonToDb(JsonElement element) => element.ValueKind switch
    {
        JsonValueKind.Null or JsonValueKind.Undefined => null,
        JsonValueKind.String => element.GetString(),
        JsonValueKind.True => 1,
        JsonValueKind.False => 0,
        JsonValueKind.Number => element.TryGetInt64(out var integer) ? integer : element.GetDouble(),
        _ => element.GetRawText(),
    };

    /// <summary>Читает все строки запроса как словари «колонка → значение».</summary>
    internal static List<Dictionary<string, object?>> SelectRows(
        SqliteConnection connection,
        string sql,
        Action<SqliteCommand>? prepare = null)
    {
        var rows = new List<Dictionary<string, object?>>();
        using var command = connection.CreateCommand();
        command.CommandText = sql;
        prepare?.Invoke(command);
        using var reader = command.ExecuteReader();
        var columnCount = reader.FieldCount;
        while (reader.Read())
        {
            var row = new Dictionary<string, object?>(columnCount);
            for (var i = 0; i < columnCount; i++)
            {
                row[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
            }
            rows.Add(row);
        }
        return rows;
    }

    private static void ExecuteDelete(SqliteConnection connection, string sql, long dealId)
    {
        using var command = connection.CreateCommand();
        command.CommandText = sql;
        Add(command, "@deal_id", dealId);
        command.ExecuteNonQuery();
    }

    /// <summary>Аналог New-PurchaseTrashItem без файлов: строки сохраняются в payload_json.</summary>
    private static void InsertTrashItem(
        SqliteConnection connection,
        string entityType,
        long entityId,
        string title,
        Dictionary<string, object> payload)
        => InsertTrashItem(connection, entityType, entityId, title, payload, null, string.Empty);

    /// <summary>
    /// Порт New-PurchaseTrashItem (PurchaseStore.ps1, строки 703-747) с файлами и папкой задачи:
    /// запись корзины создаётся первой, затем файлы переезжают в {корень корзины}/{id}
    /// под именами file_{0000N}_{имя}, папка задачи — в component_folder, после чего
    /// files_json обновляется. База не трогается вызывающим кодом до успеха переноса.
    /// </summary>
    internal static long InsertTrashItem(
        SqliteConnection connection,
        string entityType,
        long entityId,
        string title,
        Dictionary<string, object> payload,
        IReadOnlyList<string>? filePaths,
        string componentFolder)
    {
        using (var insert = connection.CreateCommand())
        {
            insert.CommandText = @"
INSERT INTO trash_items(entity_type, entity_id, title, payload_json, files_json, deleted_at)
VALUES(@entity_type, @entity_id, @title, @payload_json, '', @deleted_at);";
            Add(insert, "@entity_type", entityType);
            Add(insert, "@entity_id", entityId);
            Add(insert, "@title", title);
            Add(insert, "@payload_json", JsonSerializer.Serialize(payload));
            Add(insert, "@deleted_at", NowText());
            insert.ExecuteNonQuery();
        }

        long trashId;
        using (var selectId = connection.CreateCommand())
        {
            selectId.CommandText = "SELECT last_insert_rowid();";
            trashId = Convert.ToInt64(selectId.ExecuteScalar());
        }
        var itemRoot = Path.Combine(AppPaths.PurchaseDataDirectory, "trash_files", trashId.ToString());
        var fileRecords = new List<Dictionary<string, object>>();
        try
        {
            var number = 0;
            foreach (var filePath in (filePaths ?? Array.Empty<string>())
                         .Where(path => !string.IsNullOrWhiteSpace(path))
                         .Distinct())
            {
                if (!File.Exists(filePath))
                {
                    continue;
                }

                Directory.CreateDirectory(itemRoot);
                number++;
                var target = Path.Combine(itemRoot, $"file_{number:D4}_{Path.GetFileName(filePath)}");
                File.Move(filePath, target, overwrite: true);
                fileRecords.Add(new Dictionary<string, object>
                {
                    ["OriginalPath"] = filePath,
                    ["TrashPath"] = target,
                    ["IsFolder"] = false,
                });
            }

            if (!string.IsNullOrWhiteSpace(componentFolder) && Directory.Exists(componentFolder))
            {
                Directory.CreateDirectory(itemRoot);
                var folderTarget = Path.Combine(itemRoot, "component_folder");
                Directory.Move(componentFolder, folderTarget);
                fileRecords.Add(new Dictionary<string, object>
                {
                    ["OriginalPath"] = componentFolder,
                    ["TrashPath"] = folderTarget,
                    ["IsFolder"] = true,
                });
            }

            using var update = connection.CreateCommand();
            update.CommandText = "UPDATE trash_items SET files_json = @files_json WHERE id = @id;";
            Add(update, "@files_json", JsonSerializer.Serialize(fileRecords));
            Add(update, "@id", trashId);
            update.ExecuteNonQuery();
            return trashId;
        }
        catch (Exception exception) when (exception is not SqliteException)
        {
            throw new InvalidOperationException(
                "Не удалось переместить данные в корзину: " + exception.Message, exception);
        }
    }

    /// <summary>Относительные пути портативной сборки разрешаются от корня приложения.</summary>
    internal static string ResolvePortable(string? path)
    {
        var trimmed = path?.Trim() ?? string.Empty;
        if (trimmed.Length == 0)
        {
            return string.Empty;
        }
        return Path.IsPathRooted(trimmed) ? trimmed : Path.Combine(AppPaths.AppRoot, trimmed);
    }

    internal static long ToLong(object? value)
        => value is null or DBNull ? 0 : Convert.ToInt64(value);

    private static long GetDealId(SqliteConnection connection, string dealNumber)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT id FROM deals WHERE deal_number = @deal_number;";
        Add(command, "@deal_number", dealNumber);
        var value = command.ExecuteScalar();
        if (value is null or DBNull)
        {
            throw new InvalidOperationException("Не удалось определить созданную сделку.");
        }
        return Convert.ToInt64(value);
    }

    private static string? GetString(SqliteConnection connection, string sql, long id)
    {
        using var command = connection.CreateCommand();
        command.CommandText = sql;
        Add(command, "@id", id);
        var value = command.ExecuteScalar();
        return value is null or DBNull ? null : Convert.ToString(value);
    }

    private static string? NormalizeBoardCount(string? value)
    {
        if (value is null) return null;
        var text = value.Trim();
        if (text.Length == 0) return string.Empty;
        if (!int.TryParse(text, out var count) || count < 0)
            throw new ArgumentException("В поле «Кол-во плат» укажите целое неотрицательное число.");
        return count.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private static void TouchDeal(SqliteConnection connection, long dealId)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "UPDATE deals SET updated_at = @updated_at WHERE id = @id;";
        Add(command, "@updated_at", NowText());
        Add(command, "@id", dealId);
        command.ExecuteNonQuery();
    }

    internal static void WriteActivity(
        SqliteConnection connection,
        string entityType,
        long entityId,
        string action,
        string details,
        long dealId,
        long supplierId = 0)
    {
        // Как и оригинальный Write-ActivityLog, журнал не должен откатывать
        // полезное действие, если таблица старой базы ещё не содержит лога.
        try
        {
            using var command = connection.CreateCommand();
            command.CommandText = @"
INSERT INTO activity_log(entity_type, entity_id, deal_id, supplier_id, action, details, created_at)
VALUES(@entity_type, @entity_id, @deal_id, @supplier_id, @action, @details, @created_at);";
            Add(command, "@entity_type", entityType);
            Add(command, "@entity_id", entityId == 0 ? null : entityId);
            Add(command, "@deal_id", dealId == 0 ? null : dealId);
            Add(command, "@supplier_id", supplierId == 0 ? null : supplierId);
            Add(command, "@action", action);
            Add(command, "@details", details);
            Add(command, "@created_at", NowText());
            command.ExecuteNonQuery();
        }
        catch (SqliteException)
        {
            // Совместимость со старой базой без activity_log.
        }
    }

    internal static void Add(SqliteCommand command, string name, object? value)
        => command.Parameters.AddWithValue(name, value ?? DBNull.Value);

    private static string RequireDealNumber(string? value)
    {
        var result = value?.Trim() ?? string.Empty;
        if (result.Length == 0)
        {
            throw new ArgumentException("Введите номер сделки.", nameof(value));
        }
        return result;
    }

    /// <summary>Пустая «Маски» совместима с WinForms: в БД это значение 2.</summary>
    private static int NormalizeMasks(string? value)
    {
        var text = value?.Trim();
        if (string.IsNullOrWhiteSpace(text)) return 2;
        return text is "Да" or "Yes" or "True" or "1" ? 1 : 0;
    }

    private static string NormalizeStatus(string? value)
        => string.IsNullOrWhiteSpace(value) ? "RFQ" : value.Trim();

    private static string? TrimOrNull(string? value)
        => value is null ? null : value.Trim();

    private static string NowText()
        => DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
}
