using Interop.UIAutomationClient;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Input;

namespace LandocsRobot
{
    internal class Program
    {
        private static readonly Dictionary<string, string> _configValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, string> _organizationValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, string> _ticketValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static string _logFilePath = string.Empty;
        private static LogLevel _currentLogLevel = LogLevel.Info;

        #region Подключение утилит и параметры для них
        enum LogLevel
        {
            Fatal = 1,
            Error = 2,
            Warning = 3,
            Info = 4,
            Debug = 5
        }

        // Импорт функций из user32.dll
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetCursorPos(out POINT lpPoint);

        // Импорт функции из kernel32.dll
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetConsoleWindow();

        // Константы
        private const int SW_MINIMIZE = 6; // Команда для минимизации окна

        [Flags]
        private enum MouseFlags
        {
            Move = 0x0001,
            LeftDown = 0x0002,
            LeftUp = 0x0004,
            RightDown = 0x0008,
            RightUp = 0x0010,
            Absolute = 0x8000
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        #endregion

        static void Main(string[] args)
        {
            // Основная логика робота

            string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string configPath = Path.Combine(currentDirectory, "parameters.xml");
            string logDirectory = InitializeLogging();

            // Устанавливаем путь к файлу лога
            _logFilePath = Path.Combine(logDirectory, $"{DateTime.Now:yyyy-MM-dd}.log");
            Log(LogLevel.Info, "🤖 Запуск робота LandocsRobot");

            try
            {
                // Загрузка конфигураций
                if (!LoadConfig(configPath) || !LoadConfigOrganization(GetConfigValue("PathToOrganization")))
                {
                    Log(LogLevel.Error, "Ошибка при загрузке конфигурации. Завершение работы робота.");
                    return;
                }

                // Очистка старых файлов лога
                CleanOldLogs(logDirectory, int.TryParse(GetConfigValue("LogRetentionDays"), out int days) ? days : 30);

                string inputFolderPath = GetConfigValue("InputFolderPath");
                if (!Directory.Exists(inputFolderPath))
                {
                    Log(LogLevel.Error, $"Путь к папке входящих файлов [{inputFolderPath}] не существует. Завершение работы робота.");
                    return;
                }

                //Получение входных файлов 
                string[] ticketArrays = Directory.GetDirectories(inputFolderPath);
                int ticketCount = ticketArrays.Length;

                Log(LogLevel.Info, ticketCount > 0
                    ? $"Найдено {ticketCount} заяв(-ка) (-ок) для обработки."
                    : "Папка пуста. Заявок для обработки не найдено.");

                if (ticketCount == 0)
                {
                    return;
                }

                foreach (string ticket in ticketArrays)
                {
                    try
                    {
                        // Очистка переменной заявки
                        _ticketValues.Clear();
                        string numberTicket = Path.GetFileNameWithoutExtension(ticket).Trim();
                        _ticketValues["ticketFolderName"] = numberTicket.Replace("+", "");

                        Log(LogLevel.Info, $"Начинаю обработку заявки: {numberTicket}");

                        // Поиск и проверка файла заявки
                        string ticketJsonFile = GetFileSearchDirectory(ticket, "*.txt");
                        if (ticketJsonFile == null)
                        {
                            Log(LogLevel.Error, $"Файл заявки [SD<Номер Заявки>.txt] не найден в папке [{ticket}]. Пропускаю заявку.");
                            continue;
                        }

                        Log(LogLevel.Info, $"Файл заявки [{Path.GetFileName(ticketJsonFile)}] найден. Начинаю обработку.");

                        // Парсинг JSON файла
                        var resultParseJson = ParseJsonFile(ticketJsonFile);
                        Log(LogLevel.Info, $"Извлеченные данные: Номер заявки - [{resultParseJson.Title}], Тип - [{resultParseJson.FormType}], Организация - [{resultParseJson.OrgTitle}], ППУД - [{resultParseJson.ppudOrganization}]");

                        // Сохранение извлеченной информации
                        _ticketValues["ticketName"] = resultParseJson.Title;
                        _ticketValues["ticketOrg"] = resultParseJson.OrgTitle;
                        _ticketValues["ticketType"] = resultParseJson.FormType;
                        _ticketValues["ticketPpud"] = resultParseJson.ppudOrganization;

                        // Поиск папки ЭДО
                        string ticketEdoFolder = GetFoldersSearchDirectory(ticket, "ЭДО");
                        if (ticketEdoFolder == null)
                        {
                            Log(LogLevel.Warning, $"Папка [ЭДО] не найдена в [{ticket}]. Пропускаю заявку.");
                            continue;
                        }

                        string[] ticketEdoChildren = GetFilesAndFoldersFromDirectory(ticketEdoFolder);
                        if (ticketEdoChildren.Length == 0)
                        {
                            Log(LogLevel.Error, $"Папка [ЭДО] пуста. Пропускаю заявку.");
                            continue;
                        }

                        Log(LogLevel.Info, $"В папке [ЭДО] найдено {ticketEdoChildren.Length} элементов. Начинаю обработку файлов.");

                        // Создание и проверка структуры папок
                        if (!EnsureDirectoriesExist(ticketEdoFolder, "xlsx", "pdf", "zip", "error", "document"))
                        {
                            Log(LogLevel.Error, $"Ошибка при создании структуры папок в [{ticketEdoFolder}]. Пропускаю заявку.");
                            continue;
                        }

                        // Сортировка и перемещение файлов
                        var newFoldersEdoChildren = CreateFolderMoveFiles(ticketEdoFolder, ticketEdoChildren);
                        Log(LogLevel.Info, "Сортировка и перемещение файлов завершены.");

                        // Логирование содержимого папок
                        Log(LogLevel.Debug, $"xlsx: {GetFileshDirectory(newFoldersEdoChildren.XlsxFolder).Length} элементов.");
                        Log(LogLevel.Debug, $"pdf: {GetFileshDirectory(newFoldersEdoChildren.PdfFolder).Length} элементов.");
                        Log(LogLevel.Debug, $"zip: {GetFileshDirectory(newFoldersEdoChildren.ZipFolder).Length} элементов.");
                        Log(LogLevel.Debug, $"error: {GetFileshDirectory(newFoldersEdoChildren.ErrorFolder).Length} элементов.");

                        // Обработка файлов Excel
                        string[] xlsxFiles = XlsxContainsPDF(newFoldersEdoChildren.XlsxFolder, newFoldersEdoChildren.PdfFolder);
                        Log(LogLevel.Info, $"{xlsxFiles.Length} файл(-а) (-ов) на конвертацию в PDF.");

                        if (xlsxFiles.Length > 0)
                        {
                            ConvertToPdf(xlsxFiles, newFoldersEdoChildren.PdfFolder);
                            Log(LogLevel.Info, "Конвертация Excel в PDF завершена.");
                        }

                        // Сохранение пути к PDF
                        _ticketValues["pathPdf"] = newFoldersEdoChildren.PdfFolder;

                        Log(LogLevel.Info, $"Обработка заявки [{numberTicket}] завершена успешно.");
                    }
                    catch (Exception ticketEx)
                    {
                        Log(LogLevel.Error, $"Ошибка при обработке заявки [{ticket}]: {ticketEx.Message}");
                        continue;
                    }

                    //Обработка landocs
                    //Получаем списко файлов pdf для обработки 
                    string[] arrayPdfFiles = GetFilesAndFoldersFromDirectory(GetTicketValue("pathPdf"));
                    #region Начать обработку Landocs




                    foreach (string filePdf in arrayPdfFiles)
                    {
                        int index = 0;
                        var resultparseFileName = GetParseNameFile(Path.GetFileNameWithoutExtension(filePdf));
                        Log(LogLevel.Info, $"Начинаю работу по файлу: Индекс: [{index}], Файл: [{resultparseFileName}]. Всего файлов: [{arrayPdfFiles.Length}]");
                        //Получаем наименование контрагента
                        _ticketValues["CounterpartyName"] = resultparseFileName.CounterpartyName?.Trim() ?? string.Empty;
                        //Получаем номер документа
                        _ticketValues["FileNameNumber"] = resultparseFileName.Number?.Trim() ?? string.Empty;
                        //Получаем дату документа
                        _ticketValues["FileDate"] = resultparseFileName.FileDate?.Trim() ?? string.Empty;
                        //Получаем ИНН
                        _ticketValues["FileNameINN"] = resultparseFileName.INN?.Trim() ?? string.Empty;
                        //Получаем КПП документа
                        _ticketValues["FileNameKPP"] = resultparseFileName.KPP?.Trim() ?? string.Empty;
                        try
                        {
                            Log(LogLevel.Info, $"Запускаю Landocs.");

                            // Получение путей из конфигурации
                            string customFile = GetConfigValue("ConfigLandocsCustomFile");  // Путь к исходному файлу
                            string landocsProfileFolder = GetConfigValue("ConfigLandocsFolder");  // Папка назначения

                            #region Запуск LanDocs

                            IUIAutomationElement appElement = null;
                            IUIAutomationElement targetWindowCreateDoc = null;
                            IUIAutomationElement targetWindowCounterparty = null;
                            IUIAutomationElement targetWindowAgreement = null;
                            IUIAutomationElement targetWindowGetPdfFile = null;

                            IUIAutomationElement targetElementAgreementTree = null; 

                            try
                            {
                                // Перемещение пользовательского профиля Landocs
                                MoveCustomProfileLandocs(customFile, landocsProfileFolder);
                                Log(LogLevel.Info, "Профиль Landocs успешно перемещен.");

                                // Путь к приложению Landocs
                                string appLandocsPath = GetConfigValue("AppLandocsPath");

                                // Запуск приложения и ожидание окна
                                Log(LogLevel.Info, $"Запускаю приложение Landocs по пути: {appLandocsPath}");
                                appElement = LaunchAndFindWindow(appLandocsPath, "_robin_landocs (Мой LanDocs) - Избранное - LanDocs", 300);

                                if (appElement == null)
                                {
                                    Log(LogLevel.Error, "Окно Landocs не найдено. Завершаю работу.");
                                    throw new Exception("Окно Landocs не найдено.");
                                }

                                Log(LogLevel.Info, "Приложение Landocs успешно запущено и окно найдено.");

                                // Задержка на обработку интерфейса
                                Thread.Sleep(5000);
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при запуске Landocs: {ex.Message}");
                                throw;  // Пробрасываем исключение дальше
                            }
                            #endregion

                            #region Поиск вкладки "Главная"

                            // Поиск и клик по элементу "Главная" в ТабМеню
                            string xpathSettingAccount1 = "Pane[3]/Tab/TabItem[1]";
                            Log(LogLevel.Info, "Начинаю поиск вкладки [Главная] в навигационном меню...");

                            try
                            {
                                var targetElement1 = FindElementByXPath(appElement, xpathSettingAccount1, 60);

                                if (targetElement1 != null)
                                {
                                    Log(LogLevel.Info, "Вкладка [Главная] найдена. Выполняю клик.");
                                    ClickElementWithMouse(targetElement1);


                                    Log(LogLevel.Info, "Клик по вкладке [Главная] успешно выполнен.");
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Не удалось найти вкладку [Главная] в навигационном меню.");
                                    throw new Exception("Элемент не найден - вкладка [Главная] в навигационном меню.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске или клике по вкладке [Главная]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Поиск слева в меню элемента "Документы"

                            string xpathSettingDoc = "Pane[1]/Pane/Pane[1]/Pane/Pane/Button[2]";
                            Log(LogLevel.Info, "Начинаю поиск кнопки [Документы] в навигационном меню...");

                            try
                            {
                                var targetElementDoc = FindElementByXPath(appElement, xpathSettingDoc, 60);
                                if (targetElementDoc != null)
                                {
                                    Log(LogLevel.Info, $"Нашел ссылку [Документы] в левом навигационном меню");
                                    TryInvokeElement(targetElementDoc);
                                    Log(LogLevel.Info, "Клик по элементу [Документы] успешно выполнен.");
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Не удалось найти вкладку [Главная] в навигационном меню.");
                                    throw new Exception("Элемент не найден - элемент [Документы] в навигационном меню.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске или клике по элементу [Документы]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Клик по элементу "Документы"
                            try
                            {
                                Log(LogLevel.Info, "Нажимаем Ctrl+F для вызова окна поиска ППУД.");
                                SendKeys.SendWait("^{f}");
                                Thread.Sleep(3000);

                                // Попытка получить элемент, который сейчас в фокусе
                                var targetElementSearch = GetFocusedElement();

                                // Значение ППУД из данных заявки
                                string ppudValue = GetTicketValue("ticketPpud");

                                if (targetElementSearch != null)
                                {
                                    Log(LogLevel.Info, "Элемент окна поиска ППУД успешно найден.");

                                    // Попытка получить паттерн ValuePattern для элемента
                                    if (targetElementSearch.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is IUIAutomationValuePattern valuePattern)
                                    {
                                        // Устанавливаем значение через ValuePattern
                                        valuePattern.SetValue(ppudValue);
                                        Log(LogLevel.Info, "Значение введено в окно поиска ППУД через ValuePattern.");
                                    }
                                    else
                                    {
                                        // Если ValuePattern недоступен, используем SendKeys
                                        SendKeys.SendWait(ppudValue);
                                        Log(LogLevel.Warning, "ValuePattern недоступен. Значение введено в окно поиска ППУД через SendKeys.");
                                    }
                                }
                                else
                                {
                                    // Если элемент не найден, бросаем исключение
                                    Log(LogLevel.Error, "Не удалось найти элемент окна поиска ППУД.");
                                    throw new Exception("Элемент окна поиска ППУД не найден.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при вводе значения в окно поиска ППУД: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск ППУД
                            try
                            {
                                Log(LogLevel.Info, "Ищу кнопку [Вниз] в окне поиска ППУД.");

                                // XPath для кнопки поиска вниз
                                string xpathSettingDown = "Pane[1]/Pane/Pane[1]/Pane/Pane/Pane/Pane/Tree/Pane/Pane/Pane/Button[3]";

                                // Поиск элемента
                                var targetElementDown = FindElementByXPath(appElement, xpathSettingDown, 60);

                                if (targetElementDown != null)
                                {
                                    // Устанавливаем фокус на элемент
                                    targetElementDown.SetFocus();
                                    Log(LogLevel.Info, "Фокус успешно установлен на кнопку [Вниз].");

                                    // Даем интерфейсу время для обработки фокуса
                                    Thread.Sleep(2000);

                                    Log(LogLevel.Info, "Нажали кнопку [Вниз] в окне поиска ППУД.");
                                    TryInvokeElement(targetElementDown);
                                    Log(LogLevel.Info, "Нажали кнопку [Вниз] в окне поиска ППУД успешно.");
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Элемент кнопки [Вниз] в окне поиска ППУД не найден.");
                                    throw new Exception("Элемент кнопки [Вниз] в окне поиска ППУД не найден.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске или клике по кнопке [Вниз]: {ex.Message}");
                                throw;
                            }
                            #endregion
                            Thread.Sleep(2000);
                            #region Поиск элемента ППУД в списке Документов
                            try
                            {
                                Log(LogLevel.Info, "Начинаю поиск элемента ППУД в списке документов.");

                                // XPath для группы элементов ППУД
                                string xpathSettingItem = "Pane[1]/Pane/Pane[1]/Pane/Pane/Pane/Pane/Tree/Group";

                                // Поиск элемента группы
                                IUIAutomationElement targetElementItem = FindElementByXPath(appElement, xpathSettingItem, 60);

                                // Значение ППУД из данных заявки
                                string ppudElement = GetTicketValue("ticketPpud");

                                if (targetElementItem != null)
                                {
                                    Log(LogLevel.Info, $"Группа элементов найдена. Ищу ППУД с значением: [{ppudElement}].");

                                    // Получение всех дочерних элементов
                                    IUIAutomationElementArray children = targetElementItem.FindAll(TreeScope.TreeScope_Children, new CUIAutomation().CreateTrueCondition());

                                    if (children != null && children.Length > 0)
                                    {
                                        bool isFound = false;

                                        for (int i = 0; i < children.Length; i++)
                                        {
                                            IUIAutomationElement item = children.GetElement(i);

                                            // Получение текстового значения элемента
                                            string value = item.GetCurrentPropertyValue(UIA_PropertyIds.UIA_ValueValuePropertyId)?.ToString() ?? "Нет значения";

                                            if (value == ppudElement)
                                            {
                                                // Вызов действия для найденного элемента
                                                try
                                                {
                                                    TryInvokeElement(item);
                                                    Log(LogLevel.Info, $"ППУД [{ppudElement}] найден и успешно обработан.");
                                                    isFound = true;
                                                    break;
                                                }
                                                catch
                                                {
                                                    Log(LogLevel.Error, $"Не удалось выполнить действие для ППУД [{ppudElement}].");
                                                    throw new Exception($"Ошибка: Не удалось выполнить действие для ППУД [{ppudElement}].");
                                                }
                                            }
                                        }

                                        if (!isFound)
                                        {
                                            Log(LogLevel.Error, $"ППУД [{ppudElement}] не найден в списке.");
                                            throw new Exception($"Ошибка: ППУД [{ppudElement}] не найден в списке документов.");
                                        }
                                    }
                                    else
                                    {
                                        Log(LogLevel.Error, "Список элементов ППУД пуст или недоступен.");
                                        throw new Exception("Ошибка: Список элементов ППУД пуст или недоступен.");
                                    }
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Группа с элементами ППУД не найдена.");
                                    throw new Exception("Ошибка: Группа с элементами ППУД не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске элемента ППУД: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Нажимаем кнопку "Создать документ"
                            try
                            {
                                Log(LogLevel.Info, "Начинаю поиск кнопки для создания документа.");

                                // XPath для кнопки
                                string xpathCreateDocButton = "Pane[3]/Pane/Pane/ToolBar[1]/Button";

                                // Поиск кнопки
                                var targetElementCreateDocButton = FindElementByXPath(appElement, xpathCreateDocButton, 60);

                                if (targetElementCreateDocButton != null)
                                {
                                    // Получение имени кнопки
                                    string elementValue = targetElementCreateDocButton.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId)?.ToString() ?? "Неизвестная кнопка";

                                    Log(LogLevel.Info, $"Кнопка [{elementValue}] найдена. Устанавливаю фокус и выполняю действие.");

                                    ClickElementWithMouse(targetElementCreateDocButton);
                                    Log(LogLevel.Info, $"Успешно нажали на кнопку [{elementValue}].");
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Кнопка для создания документа не найдена.");
                                    throw new Exception("Ошибка: Кнопка для создания документа не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при нажатии на кнопку создания документа: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Окно "Создать документ"
                            try
                            {
                                Log(LogLevel.Info, "Начинаю поиск окна создания документа.");

                                string findNameWindow = "Без имени - Документ LanDocs";
                                targetWindowCreateDoc = FindElementByName(appElement, findNameWindow, 300);

                                string elementValue = null;

                                // Проверяем, был ли найден элемент
                                if (targetWindowCreateDoc != null)
                                {
                                    // Получаем значение свойства Name
                                    elementValue = targetWindowCreateDoc.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId)?.ToString();


                                    // Проверяем, соответствует ли свойство Name ожидаемому значению
                                    if (elementValue == findNameWindow)
                                    {
                                        Log(LogLevel.Info, $"Появилось окно создания документа: [{elementValue}].");
                                    }
                                    else
                                    {
                                        Log(LogLevel.Error, $"Ожидалось окно с названием 'Без имени - Документ LanDocs', но найдено: [{elementValue ?? "Неизвестное имя"}].");
                                        throw new Exception($"Неверное окно: [{elementValue ?? "Неизвестное имя"}].");
                                    }
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Окно создания документа не найдено.");
                                    throw new Exception("Окно создания документа не найдено.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске окна создания документа: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Выпадающий список "Тип документа"
                            try
                            {
                                Log(LogLevel.Info, "Начинаю процесс выбора типа документа.");

                                // XPath для комбобокса и кнопки
                                string xpathElementTypeDoc = "Tab/Pane/Pane/Pane/Tab/Pane/Pane[4]/Pane/Pane[1]/Pane[2]/Pane[14]/ComboBox";
                                string xpathButtonTypeDoc = "Button[1]";
                                string typeDocument = "ППУД. Исходящий электронный документ";

                                // Поиск комбобокса
                                var targetElementTypeDoc = FindElementByXPath(targetWindowCreateDoc, xpathElementTypeDoc, 60);

                                if (targetElementTypeDoc != null)
                                {
                                    // Поиск кнопки внутри комбобокса
                                    var targetElementTypeDocButton = FindElementByXPath(targetElementTypeDoc, xpathButtonTypeDoc, 60);

                                    if (targetElementTypeDocButton != null)
                                    {
                                        // Фокус и клик по кнопке комбобокса
                                        targetElementTypeDocButton.SetFocus();
                                        TryInvokeElement(targetElementTypeDocButton);
                                        Log(LogLevel.Info, "Открыли список выбора типа документа.");

                                        // Поиск элемента типа документа по имени
                                        var docV = FindElementByName(targetWindowCreateDoc, typeDocument, 60);
                                        if (docV != null)
                                        {
                                            TryInvokeElement(docV);
                                            Log(LogLevel.Info, $"Выбрали тип документа: [{typeDocument}].");
                                        }
                                        else
                                        {
                                            Log(LogLevel.Error, $"Элемент с именем '[{typeDocument}]' не найден.");
                                            throw new Exception($"Элемент с именем '[{typeDocument}]' не найден.");
                                        }
                                    }
                                    else
                                    {
                                        Log(LogLevel.Error, "Кнопка комбобокса для выбора типа документа не найдена.");
                                        throw new Exception("Кнопка комбобокса для выбора типа документа не найдена.");
                                    }
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Комбобокс для выбора типа документа не найден.");
                                    throw new Exception("Комбобокс для выбора типа документа не найден.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при выборе типа документа: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Выпадающий список "Вид документа"
                            try
                            {
                                Log(LogLevel.Info, "Начинаю поиск и выбор типа документа для второго типа.");

                                // XPath для второго типа документа
                                string xpathElementTypeDocSecond = "Tab/Pane/Pane/Pane/Tab/Pane/Pane[4]/Pane/Pane[1]/Pane[2]/Pane[16]/ComboBox";
                                string typeDocumentSecond = "ППУД ИСХ. Акт сверки по договору / договорам";

                                // Поиск второго элемента ComboBox
                                var targetElementTypeDocSecond = FindElementByXPath(targetWindowCreateDoc, xpathElementTypeDocSecond, 60);

                                // Проверка, найден ли элемент
                                if (targetElementTypeDocSecond != null)
                                {
                                    // Поиск кнопки внутри ComboBox
                                    var targetElementTypeDocButtonSecond = FindElementByXPath(targetElementTypeDocSecond, "Button[1]", 60);

                                    if (targetElementTypeDocButtonSecond != null)
                                    {
                                        targetElementTypeDocButtonSecond.SetFocus();
                                        TryInvokeElement(targetElementTypeDocButtonSecond);
                                        Log(LogLevel.Info, "Нажали на кнопку выбора типа документа.");

                                        // Поиск и выбор второго типа документа по имени
                                        var docVSecond = FindElementByName(targetWindowCreateDoc, typeDocumentSecond, 60);
                                        if (docVSecond != null)
                                        {
                                            TryInvokeElement(docVSecond);
                                            Log(LogLevel.Info, $"Выбрали тип документа: [{typeDocumentSecond}].");
                                        }
                                        else
                                        {
                                            Log(LogLevel.Error, $"Элемент с именем '[{typeDocumentSecond}]' не найден.");
                                            throw new Exception($"Элемент с именем '[{typeDocumentSecond}]' не найден.");
                                        }
                                    }
                                    else
                                    {
                                        Log(LogLevel.Error, "Не удалось найти кнопку внутри ComboBox для второго типа документа.");
                                        throw new Exception("Не удалось найти кнопку внутри ComboBox для второго типа документа.");
                                    }
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Не удалось найти ComboBox для второго типа документа.");
                                    throw new Exception("Не удалось найти ComboBox для второго типа документа.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при выборе типа документа для второго типа: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Кнопка открытия списка контрагентов
                            try
                            {
                                Log(LogLevel.Info, "Начинаю поиск кнопки для открытия окна с контрагентами.");

                                // XPath для кнопки "Открыть окно с контрагентами"
                                string xpathCounterpartyDocButton = "Tab/Pane/Pane/Pane/Tab/Pane/Pane[4]/Pane/Pane[1]/Pane[2]/Pane[7]/Edit/Button[1]";
                                var targetElementCounterpartyDocButton = FindElementByXPath(targetWindowCreateDoc, xpathCounterpartyDocButton, 60);

                                // Проверка, найден ли элемент
                                if (targetElementCounterpartyDocButton != null)
                                {
                                    // Попытка взаимодействия с кнопкой
                                    ClickElementWithMouse(targetElementCounterpartyDocButton);
                                    Log(LogLevel.Info, "Нажали на кнопку [Открыть окно с контрагентами].");
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Кнопка [Открыть окно с контрагентами] не найдена.");
                                    throw new Exception("Кнопка [Открыть окно с контрагентами] не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при нажатии на кнопку [Открыть окно с контрагентами]: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск окна с контрагентами
                            try
                            {
                                // Попытка найти окно по имени
                                targetWindowCounterparty = FindElementByName(targetWindowCreateDoc, "Выбор элемента", 60);

                                // Если окно не найдено, пробуем найти его по XPath
                                if (targetWindowCounterparty == null)
                                {
                                    Log(LogLevel.Warning, "Окно не найдено по имени. Пробуем найти его по XPath...");
                                    string xpathWindowCounterparty = "Window[1]";
                                    targetWindowCounterparty = FindElementByXPath(targetWindowCreateDoc, xpathWindowCounterparty, 60);
                                }

                                // Проверка, найдено ли окно
                                if (targetWindowCounterparty != null)
                                {
                                    Log(LogLevel.Info, $"Появилось окно поиска контрагента [Выбор элемента]");
                                }
                                else
                                {
                                    throw new Exception($"Окно поиска контрагента [Выбор элемента] не найдено.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске окна поиска контрагента: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Ищем элемент ввода контрагента
                            try
                            {
                                string xpatElementCounterpartyInput = "Pane[1]/Pane/Table/Pane/Pane/Edit/Edit[1]";
                                var targetElementCounterpartyInput = FindElementByXPath(targetWindowCounterparty, xpatElementCounterpartyInput, 60);

                                string counterparty = GetTicketValue("FileNameINN");

                                if (targetElementCounterpartyInput != null)
                                {
                                    // Проверяем, поддерживает ли элемент ValuePattern
                                    var valuePattern = targetElementCounterpartyInput.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) as IUIAutomationValuePattern;

                                    if (valuePattern != null)
                                    {
                                        valuePattern.SetValue(counterparty);
                                        Log(LogLevel.Info, $"Значение [{counterparty}] успешно введено в поле поиска контрагента через ValuePattern.");
                                    }
                                    else
                                    {
                                        // Если ValuePattern не поддерживается, используем SendKeys
                                        targetElementCounterpartyInput.SetFocus();
                                        SendKeys.SendWait(counterparty);
                                        Log(LogLevel.Info, $"Значение [{counterparty}] введено в поле поиска контрагента с помощью SendKeys.");
                                    }
                                }
                                else
                                {
                                    throw new Exception($"Поле поиска контрагента не найдено. Значение [{counterparty}] не удалось ввести.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при вводе значения в поле поиска контрагента: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Ищем элемент кнопка "Поиск" контрагента
                            try
                            {
                                string xpathSearchCounterpartyButton = "Pane[1]/Pane/Table/Pane/Pane/Button[2]";
                                var targetElementSearchCounterpartyButton = FindElementByXPath(targetWindowCounterparty, xpathSearchCounterpartyButton, 60);

                                if (targetElementSearchCounterpartyButton != null)
                                {
                                    var elementValue = targetElementSearchCounterpartyButton.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId)?.ToString();
                                    if (elementValue != null)
                                    {
                                        targetElementSearchCounterpartyButton.SetFocus();
                                        TryInvokeElement(targetElementSearchCounterpartyButton);
                                        Log(LogLevel.Info, $"Нажали на кнопку поиска контрагента[{elementValue}]");
                                    }
                                    else
                                    {
                                        // Если ValuePattern не поддерживается, используем SendKeys
                                        targetElementSearchCounterpartyButton.SetFocus();
                                        SendKeys.SendWait("{Enter}");
                                        Log(LogLevel.Info, $"Нажали на кнопку поиска контрагента с помощью SendKeys.");
                                    }
                                    
                                }
                                else
                                {
                                    throw new Exception($"Элемент кнопки поиcка контрагента не найден.");
                                }
                            }
                            catch(Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске элемента [Поиск] или клика по нему: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск контрагента в списке
                            try
                            {
                                // Поиск элемента ППУД в списке документов
                                string xpathCounterpartyList = "Pane[1]/Pane/Table";
                                Log(LogLevel.Info, "Начинаем поиск элемента 'Список контрагентов'...");

                                IUIAutomationElement targetElementCounterpartyList = FindElementByXPath(targetWindowCounterparty, xpathCounterpartyList, 60);
                                if (targetElementCounterpartyList == null)
                                {
                                    throw new Exception("Ошибка: Элемент 'Список контрагентов' не найден. Работа робота завершена.");
                                }

                                Log(LogLevel.Info, "Элемент 'Список контрагентов' найден. Пытаемся найти 'Панель данных' внутри списка...");
                                IUIAutomationElement dataPanel = FindElementByName(targetElementCounterpartyList, "Панель данных", 60);

                                if (dataPanel == null)
                                {
                                    throw new Exception("Ошибка: Элемент 'Панель данных' не найден. Работа робота завершена.");
                                }

                                Log(LogLevel.Info, "'Панель данных' найдена. Получаем список контрагентов...");
                                IUIAutomationElementArray childrenCounterparty = dataPanel.FindAll(
                                    TreeScope.TreeScope_Children,
                                    new CUIAutomation().CreateTrueCondition()
                                );

                                if (childrenCounterparty == null || childrenCounterparty.Length == 0)
                                {
                                    Log(LogLevel.Warning, "Список контрагентов пуст или не найден.");
                                    throw new Exception("Ошибка: Список контрагентов пуст или не найден. Работа робота завершена.");
                                }

                                Log(LogLevel.Info, $"Получен список контрагентов: найдено {childrenCounterparty.Length} элементов.");
                                var counterpartyElements = new Dictionary<int, string[]>();

                                string innValue = GetTicketValue("FileNameINN");
                                string kppValue = GetTicketValue("FileNameKPP");
                                string counterpartyName = GetTicketValue("CounterpartyName");

                                for (int i = 0; i < childrenCounterparty.Length; i++)
                                {
                                    Log(LogLevel.Debug, $"Обработка контрагента под индексом [{i}]...");

                                    IUIAutomationElement itemCounterparty = childrenCounterparty.GetElement(i);
                                    IUIAutomationElement dataItem = FindElementByXPath(itemCounterparty, "dataitem", 60);

                                    if (dataItem == null)
                                    {
                                        Log(LogLevel.Warning, $"Контрагент под индексом [{i}] не содержит элемента 'dataitem'. Пропускаем...");
                                        continue;
                                    }

                                    if (dataItem.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is IUIAutomationValuePattern valuePattern)
                                    {
                                        string value = valuePattern.CurrentValue ?? string.Empty;
                                        Log(LogLevel.Debug, $"Найден контрагент [{i}]: [{value}]");

                                        // Обрабатываем строки и добавляем в словарь
                                        counterpartyElements[i] = value
                                            .Split(',')
                                            .Select(v => v.Trim())
                                            .ToArray();
                                    }
                                    else
                                    {
                                        Log(LogLevel.Error, $"Элемент контрагента [{i}] не поддерживает ValuePattern. Пропускаем...");
                                    }
                                }

                                Log(LogLevel.Info, "Все элементы контрагентов обработаны. Выполняем поиск необходимого контрагента...");

                                int? foundKey = FindCounterpartyKey(counterpartyElements, innValue, kppValue, counterpartyName);

                                if (foundKey.HasValue)
                                {
                                    Log(LogLevel.Info, $"Необходимый контрагент найден: ключ [{foundKey.Value}].");
                                    IUIAutomationElement requiredElement = childrenCounterparty.GetElement(foundKey.Value);
                                    IUIAutomationElement selectedDataItem = FindElementByXPath(requiredElement, "dataitem", 60);

                                    if (selectedDataItem != null)
                                    {
                                        Log(LogLevel.Info, "Выбираем найденного контрагента.");
                                        selectedDataItem.SetFocus();
                                        TryInvokeElement(selectedDataItem);
                                        Log(LogLevel.Info, "Работа с найденным элементом завершена.");
                                    }
                                    else
                                    {
                                        Log(LogLevel.Error, "Ошибка: Не удалось найти 'dataitem' у найденного контрагента.");
                                        throw new Exception("Ошибка: Не удалось выбрать найденного контрагента.");
                                    }
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Ошибка: Не удалось найти контрагента с заданными ИНН, КПП и наименованием.");
                                    throw new Exception("Ошибка: Не удалось выбрать контрагента с заданными ИНН, КПП и наименованием.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске или выборе контрагента в списке результатов: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Кнопка "Выбрать" в окне поиска контрагентов
                            try
                            {
                                string xpathCounterpartyOkButton = "Pane[2]/Button[1]";
                                Log(LogLevel.Info, "Начинаем поиск кнопки [Выбрать] в окне [Выбор элемента] со списком контрагентов...");

                                // Поиск кнопки [Выбрать]
                                var targetElementCounterpartyOkButton = FindElementByXPath(targetWindowCounterparty, xpathCounterpartyOkButton, 10);

                                if (targetElementCounterpartyOkButton != null)
                                {
                                    Log(LogLevel.Info, "Кнопка [Выбрать] найдена. Пытаемся нажать на кнопку...");

                                    // Установка фокуса на кнопку и попытка нажатия
                                    targetElementCounterpartyOkButton.SetFocus();
                                    TryInvokeElement(targetElementCounterpartyOkButton);

                                    Log(LogLevel.Info, "Нажали на кнопку [Выбрать] в окне [Выбор элемента] со списком контрагентов.");
                                }
                                else
                                {
                                    // Если кнопка не найдена
                                    throw new Exception("Кнопка [Выбрать] в окне [Выбор элемента] со списком контрагентов не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске или нажатии кнопки [Выбрать]: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Кнопка [...] для открытия окна с договорами
                            try
                            {
                                string xpathAgreementButton = "Tab/Pane/Pane/Pane/Tab/Pane/Pane[3]/Pane/Pane/Button[2]";
                                Log(LogLevel.Info, "Начинаем поиск кнопки [...] для выбора договора в окне [Создание документа]...");

                                // Поиск кнопки выбора договора
                                var targetElementAgreementButton = FindElementByXPath(targetWindowCreateDoc, xpathAgreementButton, 10);

                                if (targetElementAgreementButton != null)
                                {
                                    Log(LogLevel.Info, "Кнопка открытия окна с  договорами найдена. Пытаемся нажать на кнопку...");

                                    // Установка фокуса на кнопку и попытка нажатия
                                    targetElementAgreementButton.SetFocus();
                                    ClickElementWithMouse(targetElementAgreementButton);

                                    Log(LogLevel.Info, "Нажали на кнопку открытия окна с договорами в окне [Создание документа].");
                                }
                                else
                                {
                                    // Если кнопка не найдена
                                    throw new Exception("Кнопка для выбора договора [...] в окне [Создание документа] не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске или нажатии кнопки [...] для выбора договора: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск окна с договорами [Выбор документа]
                            try
                            {
                                // Поиск окна "Выбор документа"
                                targetWindowAgreement = FindElementByName(targetWindowCreateDoc, "Выбор документа", 60);

                                // Проверка, был ли найден элемент
                                if (targetWindowAgreement != null)
                                {
                                    Log(LogLevel.Info, "Окно поиска контрагента [Выбор документа] найдено.");
                                }
                                else
                                {
                                    // Если элемент не найден
                                    throw new Exception("Ошибка: Окно поиска контрагента [Выбор документа] не найдено.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку, если возникла проблема при поиске окна
                                Log(LogLevel.Error, $"Ошибка при поиске окна [Выбор документа]: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск дерева [Журналы регистрации] в окне [Выбор документа]
                            try
                            {
                                // Поиск дерева элементов в списке [Журналы регистрации]
                                string xpathAgreementTree = "Pane/Pane/Pane[3]/Tree";
                                Log(LogLevel.Info, "Начинаем поиск дерева элементов списка [Журналы регистрации]...");

                                // Ищем элемент дерева
                                targetElementAgreementTree = FindElementByXPath(targetWindowAgreement, xpathAgreementTree, 60);

                                if (targetElementAgreementTree != null)
                                {
                                    Log(LogLevel.Info, "Элемент дерева [Журналы регистрации] найден.");
                                }
                                else
                                {
                                    // Если элемент не найден
                                    throw new Exception("Ошибка: Элемент дерева [Журналы регистрации] не найден. Работа робота завершена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку, если возникла проблема при поиске элемента
                                Log(LogLevel.Error, $"Ошибка при поиске элемента дерева [Журналы регистрации]: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск скрола в дереве [Журнала регистрации]
                            try
                            {
                                Log(LogLevel.Info, "Элемент дерева [Журналы регистрации] найден. Пытаемся инициализировать скролл...");

                                // Поиск элемента скролла
                                var targetElementAgreemenScrollBar = FindElementByName(targetElementAgreementTree, "Vertical", 60);

                                if (targetElementAgreemenScrollBar != null)
                                {
                                    // Если элемент скролла найден
                                    Log(LogLevel.Info, "Элемент скролла [Vertical] найден! Работа робота продолжается.");
                                }
                                else
                                {
                                    // Если элемент скролла не найден
                                    throw new Exception("Ошибка: Элемент скролла [Vertical] не найден! Работа робота завершена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку, если возникла проблема при поиске скролла
                                Log(LogLevel.Error, $"Ошибка при поиске скролла [Vertical]: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Проверка состояния элемента [Журналы регистраций]
                            try
                            {
                                // Поиск элемента [Журналы регистрации] в дереве
                                var targetElementAgreemenTreeItem = FindElementByName(targetElementAgreementTree, "Журналы регистрации", 60);

                                if (targetElementAgreemenTreeItem != null)
                                {
                                    Log(LogLevel.Info, "Элемент [Журналы регистраций] найден.");

                                    // Проверка, поддерживает ли элемент ExpandCollapsePattern
                                    if (targetElementAgreemenTreeItem.GetCurrentPattern(UIA_PatternIds.UIA_ExpandCollapsePatternId) is IUIAutomationExpandCollapsePattern expandCollapsePattern)
                                    {
                                        var state = expandCollapsePattern.CurrentExpandCollapseState;

                                        switch (state)
                                        {
                                            case ExpandCollapseState.ExpandCollapseState_Collapsed:
                                                Log(LogLevel.Debug, "Элемент [Журналы регистраций] свернут. Раскрываем...");
                                                expandCollapsePattern.Expand(); // Раскрываем элемент
                                                Log(LogLevel.Info, "Элемент [Журналы регистраций] успешно раскрыт.");
                                                break;

                                            case ExpandCollapseState.ExpandCollapseState_Expanded:
                                                Log(LogLevel.Debug, "Элемент [Журналы регистраций] уже раскрыт.");
                                                break;

                                            case ExpandCollapseState.ExpandCollapseState_PartiallyExpanded:
                                                Log(LogLevel.Debug, "Элемент [Журналы регистраций] частично раскрыт. Раскрываем полностью...");
                                                expandCollapsePattern.Expand(); // Раскрываем элемент
                                                break;

                                            case ExpandCollapseState.ExpandCollapseState_LeafNode:
                                                Log(LogLevel.Debug, "Элемент [Журналы регистраций] является листовым узлом. Раскрытие не требуется.");
                                                break;

                                            default:
                                                Log(LogLevel.Warning, "Неизвестное состояние ExpandCollapseState.");
                                                break;
                                        }
                                    }
                                    else
                                    {
                                        throw new Exception("Элемент [Журналы регистраций] не поддерживает ExpandCollapsePattern.");
                                    }
                                }
                                else
                                {
                                    throw new Exception("Элемент [Журналы регистраций] не найден.");
                                }

                                #region Поиск договора в дереве [Журналы регистраций]
                                try
                                {
                                    // Находим дочерние элементы
                                    IUIAutomationElementArray childrenAgreemen = targetElementAgreemenTreeItem.FindAll(
                                        TreeScope.TreeScope_Children,
                                        new CUIAutomation().CreateTrueCondition()
                                    );

                                    if (childrenAgreemen != null && childrenAgreemen.Length > 0)
                                    {
                                        bool isFound = false;
                                        int count = childrenAgreemen.Length;

                                        Log(LogLevel.Info, $"Количество журналов [{count}]");

                                        for (int i = 0; i < count; i++)
                                        {
                                            var childElement = childrenAgreemen.GetElement(i);

                                            if (childElement != null)
                                            {
                                                // Проверяем наличие LegacyIAccessiblePattern
                                                if (childElement.GetCurrentPattern(UIA_PatternIds.UIA_LegacyIAccessiblePatternId) is IUIAutomationLegacyIAccessiblePattern legacyPattern)
                                                {
                                                    string name = legacyPattern.CurrentName;

                                                    string agreementName = GetTicketValue("ticketPpud");
                                                    var agreementNameSplit = agreementName.Split('.')[0]; // Возьмем часть строки до первой точки
                                                    var agreementNameFull = string.Concat(agreementNameSplit, ".", "Договоры").ToString();
                                                    var agreementNameNormalize = agreementNameFull.Trim().ToLower().Replace(" ", "");

                                                    Log(LogLevel.Debug, $"Выполняю поиск внутри элемента [Журналы регистраций] - Журнал [{agreementNameFull}]. Фокус на элементе: [{name}]");

                                                    // Сравниваем имена с нормализацией
                                                    if (agreementNameNormalize == name.Trim().ToLower().Replace(" ", ""))
                                                    {
                                                        Log(LogLevel.Info, $"Журнал [{agreementNameFull}] внутри элемента [Журналы регистраций] найден.");

                                                        // Прокручиваем до элемента, если поддерживается ScrollItemPattern
                                                        if (childElement.GetCurrentPattern(UIA_PatternIds.UIA_ScrollItemPatternId) is IUIAutomationScrollItemPattern scrollItemPattern)
                                                        {
                                                            scrollItemPattern.ScrollIntoView();
                                                            Log(LogLevel.Debug, "Элемент журнала прокручен в область видимости.");
                                                            Thread.Sleep(500);
                                                        }

                                                        // Выбираем элемент, если поддерживается SelectionItemPattern
                                                        if (childElement.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId) is IUIAutomationSelectionItemPattern selectionItemPattern)
                                                        {
                                                            childElement.SetFocus();
                                                            selectionItemPattern.Select();
                                                            Log(LogLevel.Info, "Элемент журнала выбран.");
                                                        }

                                                        isFound = true;
                                                        break;
                                                    }
                                                }
                                            }
                                        }

                                        // Если элемент не найден, прокручиваем вниз и повторяем поиск
                                        if (!isFound)
                                        {
                                            Log(LogLevel.Debug, "Элемент не найден. Прокручиваем вниз.");
                                            var scrollPattern = targetElementAgreemenTreeItem.GetCurrentPattern(UIA_PatternIds.UIA_ScrollPatternId) as IUIAutomationScrollPattern;

                                            if (scrollPattern != null && scrollPattern.CurrentVerticallyScrollable != 0)
                                            {
                                                while (scrollPattern.CurrentVerticalScrollPercent < 100)
                                                {
                                                    scrollPattern.Scroll(ScrollAmount.ScrollAmount_NoAmount, ScrollAmount.ScrollAmount_LargeIncrement);
                                                    Log(LogLevel.Debug, "Прокручиваем вниз.");

                                                    // Повторяем поиск дочерних элементов после прокрутки
                                                    childrenAgreemen = targetElementAgreemenTreeItem.FindAll(
                                                        TreeScope.TreeScope_Children,
                                                        new CUIAutomation().CreateTrueCondition()
                                                    );

                                                    for (int i = 0; i < childrenAgreemen.Length; i++)
                                                    {
                                                        var childElement = childrenAgreemen.GetElement(i);

                                                        if (childElement != null &&
                                                            childElement.GetCurrentPattern(UIA_PatternIds.UIA_LegacyIAccessiblePatternId) is IUIAutomationLegacyIAccessiblePattern legacyPattern)
                                                        {
                                                            string name = legacyPattern.CurrentName;

                                                            string agreementName = GetTicketValue("ticketPpud");
                                                            var agreementNameSplit = agreementName.Split('.')[0]; // Возьмем часть строки до первой точки
                                                            var agreementNameFull = string.Concat(agreementNameSplit, ".", "Договоры").ToString();
                                                            var agreementNameNormalize = agreementNameFull.Trim().ToLower().Replace(" ", "");

                                                            Log(LogLevel.Debug, $"Выполняю поиск внутри элемента [Журналы регистраций] - Журнал [{agreementNameFull}]. Фокус на элементе: [{name}]");

                                                            if (agreementNameNormalize == name.Trim().ToLower().Replace(" ", ""))
                                                            {
                                                                Log(LogLevel.Info, $"Журнал [{agreementNameFull}] внутри элемента [Журналы регистраций] найден.");

                                                                if (childElement.GetCurrentPattern(UIA_PatternIds.UIA_ScrollItemPatternId) is IUIAutomationScrollItemPattern scrollItemPattern)
                                                                {
                                                                    scrollItemPattern.ScrollIntoView();
                                                                    Log(LogLevel.Debug, "Элемент журнала прокручен в область видимости.");
                                                                    Thread.Sleep(500);
                                                                }

                                                                if (childElement.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId) is IUIAutomationSelectionItemPattern selectionItemPattern)
                                                                {
                                                                    childElement.SetFocus();
                                                                    selectionItemPattern.Select();
                                                                    Log(LogLevel.Info, "Элемент журнала выбран.");
                                                                }

                                                                isFound = true;
                                                                break;
                                                            }
                                                        }
                                                    }

                                                    if (isFound)
                                                        break;
                                                }
                                            }

                                            if (!isFound)
                                            {
                                                throw new Exception("Журнал не найден или возникла ошибка при обработке элемента [Журналы регистраций].");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        throw new Exception("Журнал не найден или возникла ошибка при обработке элемента [Журналы регистраций].");
                                    }

                                    #region Выбор первого полученного договора
                                    try
                                    {
                                        string xpathElementAgreementTable = "Pane/Pane/Pane[2]/Pane[2]/Table";
                                        var targetElementAgreementTable = FindElementByXPath(targetWindowAgreement, xpathElementAgreementTable, 60);

                                        if (targetElementAgreementTable != null)
                                        {
                                            // Поиск элемента "Панель данных"
                                            var targetElementAgreementTableList = FindElementByName(targetElementAgreementTable, "Панель данных", 60);

                                            if (targetElementAgreementTableList != null)
                                            {
                                                // Поиск первого дочернего элемента
                                                var automation = new CUIAutomation();
                                                IUIAutomationElement childrenAgreementTable = targetElementAgreementTableList.FindFirst(
                                                    TreeScope.TreeScope_Children,
                                                    automation.CreateTrueCondition()
                                                );

                                                if (childrenAgreementTable != null)
                                                {
                                                    try
                                                    {
                                                        // Попытка использовать LegacyIAccessiblePattern для установки фокуса
                                                        if (childrenAgreementTable.GetCurrentPattern(UIA_PatternIds.UIA_LegacyIAccessiblePatternId) is IUIAutomationLegacyIAccessiblePattern legacyPattern)
                                                        {
                                                            // Устанавливаем фокус и выбираем элемент
                                                            legacyPattern.Select((int)AccessibleSelection.TakeSelection);
                                                            legacyPattern.Select((int)AccessibleSelection.TakeFocus);
                                                            Log(LogLevel.Info, "Элемент найден. Фокус успешно установлен на первый элемент таблицы договоров.");
                                                        }
                                                        else
                                                        {
                                                            throw new Exception("LegacyIAccessiblePattern не поддерживается для элемента.");
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        throw new Exception("Не удалось выбрать первый договор. Работа робота завершена.", ex);
                                                    }
                                                }
                                                else
                                                {
                                                    throw new Exception("Список договоров пуст. Работа робота завершена, проверьте список договоров.");
                                                }
                                            }
                                            else
                                            {
                                                throw new Exception("Элемент 'Панель данных' не найден. Работа робота завершена.");
                                            }
                                        }
                                        else
                                        {
                                            throw new Exception("Таблица договоров не найдена. Работа робота завершена.");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                        throw;
                                    }
                                    #endregion
                                }
                                catch (Exception ex)
                                {
                                    Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                    throw;
                                }
                                #endregion
                                #region Поиск и нажатие кнопки "Выбрать"
                                try
                                {
                                    string xpathAgreementOkButton = "Pane/Pane/Pane[2]/Pane[3]/Button[1]";
                                    var targetElementAgreementOkButton = FindElementByXPath(targetWindowAgreement, xpathAgreementOkButton, 60);

                                    if (targetElementAgreementOkButton != null)
                                    {
                                        // Устанавливаем фокус на кнопку и нажимаем
                                        targetElementAgreementOkButton.SetFocus();
                                        TryInvokeElement(targetElementAgreementOkButton);
                                        Log(LogLevel.Info, "Нажали на кнопку [Выбрать] в окне [Выбор документа] со списком журналов.");
                                    }
                                    else
                                    {
                                        // Выбрасываем исключение, если элемент не найден
                                        throw new Exception("Кнопка [Выбрать] в окне [Выбор документа] со списком журналов не найдена.");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // Логируем ошибку и выбрасываем исключение дальше
                                    Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                    throw;
                                }
                                #endregion
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Проверяем, что договор проставлен
                            try
                            {
                                string xpathAgreementLabel = "Tab/Pane/Pane/Pane/Tab/Pane/Pane[3]/Pane/Pane/Button[4]";
                                var targetElementAgreementLabelButton = FindElementByXPath(targetWindowCreateDoc, xpathAgreementLabel, 60);

                                if (targetElementAgreementLabelButton != null)
                                {
                                    // Получаем значение свойства Name
                                    string agreementLabelName = targetElementAgreementLabelButton.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string;

                                    // Проверяем, что значение не пустое
                                    if (!string.IsNullOrEmpty(agreementLabelName))
                                    {
                                        Log(LogLevel.Info, $"Договор проставлен успешно. Номер договора: {agreementLabelName}");
                                    }
                                    else
                                    {
                                        // Если значение пустое, выбрасываем исключение
                                        throw new Exception("Договор не проставлен, проверьте корректность. Робот завершает работу.");
                                    }
                                }
                                else
                                {
                                    // Если элемент не найден, выбрасываем исключение с детализированным сообщением
                                    throw new Exception("Кнопка [Выбрать] в окне [Выбор документа] со списком журналов не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку и выбрасываем исключение дальше
                                Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск и ввод подписанта в элемент "Подписант"
                            string xpathSignerInput = "Tab/Pane/Pane/Pane/Tab/Pane/Pane[4]/Pane/Pane[1]/Pane[1]/Pane[13]/Edit";
                            var targetElementSignerInput = FindElementByXPath(targetWindowCreateDoc, xpathSignerInput, 60);

                            if (targetElementSignerInput != null)
                            {
                                string signer = GetConfigValue("Signatory").Trim(); // Получаем значение подписанта из конфигурации
                                string currentSignerInput = targetElementSignerInput.GetCurrentPropertyValue(UIA_PropertyIds.UIA_ValueValuePropertyId) as string;

                                if (!string.IsNullOrEmpty(currentSignerInput))
                                {
                                    Log(LogLevel.Info, $"Текущий подписант: [{currentSignerInput}]. Меняю на: [{signer}].");
                                }
                                else
                                {
                                    Log(LogLevel.Info, $"Текущий подписант отсутствует. Устанавливаю нового: [{signer}].");
                                }

                                try
                                {
                                    // Используем ValuePattern для установки значения
                                    if (targetElementSignerInput.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is IUIAutomationValuePattern valuePattern)
                                    {
                                        valuePattern.SetValue(signer);
                                        Log(LogLevel.Info, $"Подписант успешно установлен: [{signer}].");
                                    }
                                    else
                                    {
                                        throw new Exception("Элемент ввода подписанта не поддерживает ValuePattern.");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    throw new Exception($"Ошибка при установке подписанта: {ex.Message}", ex);
                                }
                            }
                            else
                            {
                                throw new Exception("Элемент ввода подписанта не найден. Робот завершает работу.");
                            }
                            #endregion

                            #region Поиск и нажатие кнопки "Сохранить документ"
                            try
                            {
                                string xpathAgreementOkButton = "Pane[2]/Pane/Pane/ToolBar[1]/Button[1]";
                                var targetElementAgreementOkButton = FindElementByXPath(targetWindowCreateDoc, xpathAgreementOkButton, 60);

                                if (targetElementAgreementOkButton != null)
                                {
                                    // Устанавливаем фокус на кнопку и нажимаем
                                    targetElementAgreementOkButton.SetFocus();
                                    ClickElementWithMouse(targetElementAgreementOkButton);
                                    Log(LogLevel.Info, "Нажали на кнопку [Сохранить документ] в окне [Создать документ].");
                                }
                                else
                                {
                                    // Выбрасываем исключение, если элемент не найден
                                    throw new Exception("Кнопка [Сохранить документ] в окне [Создать документ] не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку и выбрасываем исключение дальше
                                Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск и клик на вкладку "Структура папок"
                            try
                            {
                                string xpathStructurekFolderTab = "Tab/Pane/Pane/Pane/Tab";
                                var targetElementStructurekFolderTab = FindElementByXPath(targetWindowCreateDoc, xpathStructurekFolderTab, 60);

                                if (targetElementStructurekFolderTab != null)
                                {
                                    // Поиск элемента "Панель данных"
                                    var targetElementStructurekFolderItem = FindElementByName(targetElementStructurekFolderTab, "Структура папок", 60);

                                    int retryCount = 0;
                                    bool isEnabled = false;

                                    // Проверка на доступность элемента
                                    while (targetElementStructurekFolderItem != null && retryCount < 3)
                                    {
                                        isEnabled = (bool)targetElementStructurekFolderItem.GetCurrentPropertyValue(UIA_PropertyIds.UIA_IsEnabledPropertyId);

                                        if (isEnabled)
                                        {
                                            break;
                                        }

                                        Log(LogLevel.Info, "Элемент неактивен, ждем 1 минуту...");
                                        Thread.Sleep(60000); // Ждем 1 минуту
                                        targetElementStructurekFolderItem = FindElementByName(targetElementStructurekFolderTab, "Структура папок", 60); // Переходим к следующей попытке
                                        retryCount++;
                                    }

                                    if (isEnabled)
                                    {
                                        // Получаем паттерн SelectionItemPattern
                                        if (targetElementStructurekFolderItem.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId) is IUIAutomationSelectionItemPattern SelectionItemPattern)
                                        {
                                            SelectionItemPattern.Select();
                                            ClickElementWithMouse(targetElementStructurekFolderItem);
                                            Log(LogLevel.Info, "Элемент [Структура папок] выбран.");
                                        }
                                        else
                                        {
                                            // Если паттерн не доступен, выбрасываем исключение
                                            throw new Exception("Паттерн SelectionItemPattern не поддерживается для элемента [Структура папок].");
                                        }
                                    }
                                    else
                                    {
                                        // Если элемент неактивен после 3 попыток, выбрасываем исключение
                                        throw new Exception("Элемент [Структура папок] не активен после 3 попыток.");
                                    }

                                    // Устанавливаем фокус на кнопку и нажимаем
                                    targetElementStructurekFolderTab.SetFocus();
                                    //TryInvokeElement(targetElementStructurekFolderTab);
                                    Log(LogLevel.Info, "Нажали на кнопку [Структура папок] в окне [Создать документ].");
                                }
                                else
                                {
                                    // Выбрасываем исключение, если элемент не найден
                                    throw new Exception("Кнопка [Структура папок] в окне [Создать документ] не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку и выбрасываем исключение дальше
                                Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Поиск и проверка дерева "Структуры папок"
                            try
                            {
                                string xpathStructurekFolderList = "Tab/Pane/Pane/Pane/Tab/Pane/Pane/Tree";
                                var targetElementStructurekFolderTList = FindElementByXPath(targetWindowCreateDoc, xpathStructurekFolderList, 60);

                                if (targetElementStructurekFolderTList != null)
                                {
                                    // Получаем первый дочерний элемент
                                    var childrenCheckBox = targetElementStructurekFolderTList.FindFirst(TreeScope.TreeScope_Children, new CUIAutomation().CreateTrueCondition());

                                    if (childrenCheckBox != null)
                                    {
                                        // Проверяем, является ли элемент CheckBox
                                        var togglePattern = childrenCheckBox.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId) as IUIAutomationTogglePattern;

                                        if (togglePattern != null)
                                        {
                                            // Устанавливаем значение CheckBox на true, если оно не выбрано
                                            if (togglePattern.CurrentToggleState != ToggleState.ToggleState_On)
                                            {
                                                togglePattern.Toggle();
                                                Log(LogLevel.Info, "CheckBox был установлен в состояние 'true'.");
                                            }
                                            else
                                            {
                                                Log(LogLevel.Info, "CheckBox уже установлен в состояние 'true'.");
                                            }

                                            // Ждем, чтобы элемент раскрылся после взаимодействия с CheckBox
                                            Thread.Sleep(1000);

                                            // Ищем элемент "Акты сверки" после раскрытия
                                            var checkBoxElementItem = FindElementByName(targetElementStructurekFolderTList, "Акт сверки", 60);

                                            if (checkBoxElementItem != null)
                                            {
                                                // Устанавливаем фокус на элемент и активируем его
                                                checkBoxElementItem.SetFocus();
                                                //TryInvokeElement(checkBoxElementItem);
                                                Log(LogLevel.Info, "Выбран элемент 'Акты сверки' после раскрытия CheckBox.");
                                            }
                                            else
                                            {
                                                throw new Exception("Элемент 'Акты сверки' не найден.");
                                            }
                                        }
                                        else
                                        {
                                            throw new Exception("Дочерний элемент не является CheckBox или не поддерживает TogglePattern.");
                                        }
                                    }
                                    else
                                    {
                                        throw new Exception("Не удалось найти первый дочерний элемент.");
                                    }
                                }
                                else
                                {
                                    throw new Exception("Элемент [Структура папок] не найден в окне [Создать документ].");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку и выбрасываем исключение дальше
                                Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск и нажатие кнопки "Добавить"
                            try
                            {
                                // XPath для панели с кнопкой "Добавить"
                                string xpathStructurekFolderAddTab = "Tab/Pane/Pane/Pane/Tab/Pane/Pane/Pane[6]";
                                var targetElementStructurekFolderAddTab = FindElementByXPath(targetWindowCreateDoc, xpathStructurekFolderAddTab, 60);

                                if (targetElementStructurekFolderAddTab != null)
                                {
                                    // Поиск кнопки "Добавить" внутри найденной панели
                                    var targetElementStructurekFolderAddButton = FindElementByName(targetElementStructurekFolderAddTab, "Добавить", 60);

                                    if (targetElementStructurekFolderAddButton != null)
                                    {
                                        targetElementStructurekFolderAddButton.SetFocus();
                                        ClickElementWithMouse(targetElementStructurekFolderAddButton);
                                        Log(LogLevel.Info, "Нажали на кнопку [Добавить] в окне [Создать документ].");
                                    }
                                    else
                                    {
                                        throw new Exception("Кнопка [Добавить] в панели [Создать документ] не найдена.");
                                    }
                                }
                                else
                                {
                                    throw new Exception("Панель для кнопки [Добавить] в окне [Создать документ] не найдена.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку и выбрасываем исключение дальше
                                Log(LogLevel.Error, $"Ошибка при нажатии на кнопку [Добавить] в окне [Создать документ]: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Окно "Выбрать акт (документ) pdf "
                            try
                            {
                                Log(LogLevel.Info, "Начинаю поиск окна для выбора pdf файла");

                                string findNameWindow = "Выберете файлы для прикрепления к РК";
                                targetWindowGetPdfFile = FindElementByName(targetWindowCreateDoc, findNameWindow, 300);

                                string elementValue = null;

                                // Проверяем, был ли найден элемент
                                if (targetWindowGetPdfFile != null)
                                {
                                    // Получаем значение свойства Name
                                    elementValue = targetWindowGetPdfFile.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId)?.ToString();


                                    // Проверяем, соответствует ли свойство Name ожидаемому значению
                                    if (elementValue == findNameWindow)
                                    {
                                        Log(LogLevel.Info, $"Появилось окно для прикрепления к РК: [{elementValue}].");
                                    }
                                    else
                                    {
                                        Log(LogLevel.Error, $"Ожидалось окно с названием 'Выберете файлы для прикрепления к РК', но найдено: [{elementValue ?? "Неизвестное имя"}].");
                                        throw new Exception($"Неверное окно: [{elementValue ?? "Неизвестное имя"}].");
                                    }
                                }
                                else
                                {
                                    Log(LogLevel.Error, "Окно для прикрепления к РК не найдено.");
                                    throw new Exception("Окно для прикрепления к РК не найдено.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске окна создания документа: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск и ввод пути к файлу в элемент "File Name"

                            try
                            {
                                string xpathFileName = "ComboBox[1]/Edit";
                                var targetElementFileName = FindElementByXPath(targetWindowGetPdfFile, xpathFileName, 60);

                                if (targetElementFileName == null)
                                {
                                    throw new Exception("Элемент ввода пути к файлу не найден. Робот завершает работу.");
                                }

                                string pdfFileName = filePdf.Trim(); // Получаем путь к файлу
                                string currentFileName = targetElementFileName.GetCurrentPropertyValue(UIA_PropertyIds.UIA_ValueValuePropertyId) as string;

                                if (!string.IsNullOrEmpty(currentFileName))
                                {
                                    Log(LogLevel.Debug, $"Текущий путь к файлу: [{currentFileName}]. Меняю на: [{pdfFileName}].");
                                }
                                else
                                {
                                    Log(LogLevel.Debug, $"Текущий путь к файлу отсутствует. Устанавливаю новый: [{pdfFileName}].");
                                }

                                // Используем ValuePattern для установки значения
                                if (targetElementFileName.GetCurrentPattern(UIA_PatternIds.UIA_ValuePatternId) is IUIAutomationValuePattern valuePattern)
                                {
                                    valuePattern.SetValue(pdfFileName);
                                    Log(LogLevel.Info, $"Путь к файлу успешно установлен: [{pdfFileName}].");
                                }
                                else
                                {
                                    throw new Exception("Элемент ввода пути к файлу не поддерживает ValuePattern.");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку и выбрасываем исключение дальше
                                Log(LogLevel.Error, $"Ошибка при обработке элемента ввода пути к файлу: {ex.Message}");
                                throw;
                            }

                            #endregion

                            #region Ищем элемент кнопка "Поиск" контрагента
                            try
                            {
                                // Имена для поиска
                                string nameOpen = "Open";
                                string nameOpenAlternative = "Открыть";

                                // Сначала ищем элемент по имени "Open"
                                var targetElementWindowGetPdfFile = FindElementByName(targetWindowGetPdfFile, nameOpen, 60);

                                // Если не нашли, ищем по альтернативному имени "Открыть"
                                if (targetElementWindowGetPdfFile == null)
                                {
                                    Log(LogLevel.Debug, $"Элемент с именем [{nameOpen}] не найден. Попытка поиска с именем [{nameOpenAlternative}].");
                                    targetElementWindowGetPdfFile = FindElementByName(targetWindowGetPdfFile, nameOpenAlternative, 60);
                                }

                                if (targetElementWindowGetPdfFile != null)
                                {
                                    var elementValue = targetElementWindowGetPdfFile.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId)?.ToString();

                                    if (!string.IsNullOrEmpty(elementValue))
                                    {
                                        // Устанавливаем фокус и выполняем клик
                                        targetElementWindowGetPdfFile.SetFocus();
                                        TryInvokeElement(targetElementWindowGetPdfFile);
                                        Log(LogLevel.Info, $"Нажали на кнопку [{elementValue}].");
                                    }
                                }
                                else
                                {
                                    throw new Exception($"Элемент кнопки [{nameOpen}] или [{nameOpenAlternative}] не найден.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log(LogLevel.Error, $"Ошибка при поиске элемента [Open/Открыть] или клика по нему: {ex.Message}");
                                throw;
                            }
                            #endregion

                            #region Поиск и проверка дерева "Структуры папок" и проверка что файл был прикреплен
                            try
                            {
                                string xpathStructurekFolderList = "Tab/Pane/Pane/Pane/Tab/Pane/Pane/Tree";
                                var targetElementStructurekFolderTList = FindElementByXPath(targetWindowCreateDoc, xpathStructurekFolderList, 60);

                                if (targetElementStructurekFolderTList != null)
                                {
                                    // Получаем первый дочерний элемент
                                    var childrenCheckBox = targetElementStructurekFolderTList.FindFirst(TreeScope.TreeScope_Children, new CUIAutomation().CreateTrueCondition());

                                    if (childrenCheckBox != null)
                                    {
                                        // Проверяем, является ли элемент CheckBox
                                        var togglePattern = childrenCheckBox.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId) as IUIAutomationTogglePattern;

                                        if (togglePattern != null)
                                        {
                                            // Устанавливаем значение CheckBox на true, если оно не выбрано
                                            if (togglePattern.CurrentToggleState != ToggleState.ToggleState_On)
                                            {
                                                togglePattern.Toggle();
                                                Log(LogLevel.Info, "CheckBox был установлен в состояние 'true'.");
                                            }
                                            else
                                            {
                                                Log(LogLevel.Info, "CheckBox уже установлен в состояние 'true'.");
                                            }

                                            // Ждем, чтобы элемент раскрылся после взаимодействия с CheckBox
                                            Thread.Sleep(1000);

                                            // Ищем элемент "Акты сверки" после раскрытия
                                            var checkBoxElementItem = FindElementByName(targetElementStructurekFolderTList, "Акт сверки", 60);

                                            if (checkBoxElementItem != null)
                                            {
                                                //TODO: Сделать проверку что файл был прикреплен
                                                // Устанавливаем фокус на элемент и активируем его
                                                checkBoxElementItem.SetFocus();
                                                //TryInvokeElement(checkBoxElementItem);
                                                Log(LogLevel.Info, "Выбран элемент 'Акты сверки' после раскрытия CheckBox.");


                                            }
                                            else
                                            {
                                                throw new Exception("Элемент 'Акты сверки' не найден.");
                                            }
                                        }
                                        else
                                        {
                                            throw new Exception("Дочерний элемент не является CheckBox или не поддерживает TogglePattern.");
                                        }
                                    }
                                    else
                                    {
                                        throw new Exception("Не удалось найти первый дочерний элемент.");
                                    }
                                }
                                else
                                {
                                    throw new Exception("Элемент [Структура папок] не найден в окне [Создать документ].");
                                }
                            }
                            catch (Exception ex)
                            {
                                // Логируем ошибку и выбрасываем исключение дальше
                                Log(LogLevel.Error, $"Ошибка: {ex.Message}");
                                throw;
                            }
                            #endregion


                        }
                        catch (Exception landocsEx)
                        {
                            Log(LogLevel.Error, $"Ошибка в работе LanDocs [{ticket}]: {landocsEx.Message}");
                            MessageBox.Show($"Ошибка в работе LanDocs [{ticket}]: {landocsEx.Message}");
                            continue;
                        }
                    }

                    #endregion
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Fatal, $"Глобальная ошибка: {ex.Message}");
            }
            finally
            {
                Log(LogLevel.Info, "Робот завершил работу.");
            }
        }

        #region Методы

        /// <summary>
        /// Инициализация системы логирования.
        /// </summary>
        static string InitializeLogging()
        {
            string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
            if (!Directory.Exists(logDirectory))
                Directory.CreateDirectory(logDirectory);
            return logDirectory;
        }

        /// <summary>
        /// Логирование сообщений с уровнем.
        /// </summary>
        static void Log(LogLevel level, string message)
        {
            if (level > _currentLogLevel)
            {
                return;
            }

            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            string ticketFolder = GetTicketValue("ticketFolderName");
            string context = string.IsNullOrWhiteSpace(ticketFolder) ? string.Empty : $"[{ticketFolder}] ";
            string formattedMessage = $"{timestamp} [{level}] {context}{message}";

            if (!string.IsNullOrWhiteSpace(_logFilePath))
            {
                try
                {
                    File.AppendAllText(_logFilePath, formattedMessage + Environment.NewLine);
                }
                catch (IOException ex)
                {
                    Console.Error.WriteLine($"Не удалось записать сообщение в лог: {ex.Message}");
                }
            }

            Console.WriteLine(formattedMessage);
        }

        /// <summary>
        /// Загрузка параметров конфигурации.
        /// </summary>
        static bool LoadParameters(
            string filePath,
            Dictionary<string, string> targetDictionary,
            string missingFileMessage,
            string successMessage,
            string errorMessage)
        {
            if (!File.Exists(filePath))
            {
                Log(LogLevel.Error, missingFileMessage);
                return false;
            }

            try
            {
                var document = XDocument.Load(filePath);

                if (document.Root == null)
                {
                    Log(LogLevel.Error, $"Файл {filePath} не содержит корневой элемент.");
                    return false;
                }

                targetDictionary.Clear();

                foreach (var parameter in document.Root.Elements("Parameter"))
                {
                    string name = parameter.Attribute("name")?.Value;
                    string value = parameter.Attribute("value")?.Value;

                    if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(value))
                    {
                        continue;
                    }

                    targetDictionary[name] = value;
                }

                Log(LogLevel.Info, successMessage);
                return true;
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"{errorMessage}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Загрузка параметров конфигурации.
        /// </summary>
        static bool LoadConfig(string configPath)
        {
            if (!LoadParameters(
                    configPath,
                    _configValues,
                    "Файл config.xml не найден.",
                    "Параметры успешно загружены из config.xml",
                    "Ошибка при загрузке параметров"))
            {
                return false;
            }

            string logLevelStr = GetConfigValue("LogLevel");
            if (Enum.TryParse(logLevelStr, true, out LogLevel logLevel))
            {
                _currentLogLevel = logLevel;
                Log(LogLevel.Info, $"Уровень логирования установлен на: {_currentLogLevel}");
            }
            else if (!string.IsNullOrWhiteSpace(logLevelStr))
            {
                Log(LogLevel.Warning, $"Не удалось разобрать уровень логирования '{logLevelStr}'. Используется значение по умолчанию {_currentLogLevel}.");
            }

            return true;
        }

        /// <summary>
        /// Получение значения из параметриа конфигурации.
        /// </summary>
        static string GetConfigValue(string key) => _configValues.TryGetValue(key, out var value) ? value : string.Empty;

        /// <summary>
        /// Загрузка параметров с ППУД.
        /// </summary>
        static bool LoadConfigOrganization(string pathToOrganization)
        {
            return LoadParameters(
                pathToOrganization,
                _organizationValues,
                "Не найден файл с перечислением организаций.",
                "Список организаций успешно загружен.",
                "Ошибка при загрузке списка организаций");
        }

        /// <summary>
        /// Получение значений параметров с файла с ППУД.
        /// </summary>
        static string GetConfigOrganization(string key) => _organizationValues.TryGetValue(key, out var value) ? value : string.Empty;

        /// <summary>
        /// Получение значения из текущей заявки.
        /// </summary>
        static string GetTicketValue(string key) => _ticketValues.TryGetValue(key, out var value) ? value : string.Empty;

        /// <summary>
        /// Метод очистки логов
        /// </summary>
        static void CleanOldLogs(string logDirectory, int retentionDays)
        {
            foreach (var log in Directory.EnumerateFiles(logDirectory, "*.txt").Where(f => File.GetCreationTime(f) < DateTime.Now.AddDays(-retentionDays)))
            {
                try
                {
                    File.Delete(log);
                    Log(LogLevel.Info, $"Лог-файл {log} удален");
                }
                catch (Exception e)
                {
                    Log(LogLevel.Error, $"Ошибка при удалении файла лога {log}: {e.Message}");
                }
            }
        }

        /// <summary>
        /// Получение массива с файлами и папками
        /// </summary>
        static string[] GetFilesAndFoldersFromDirectory(string folder)
        {
            try
            {
                return Directory.GetFiles(folder)
                    .Concat(Directory.GetDirectories(folder))
                    .ToArray();
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Ошибка при получении файлов и папок из папки {folder}: {ex.Message}");
                return Array.Empty<string>();  // Возвращаем пустой массив при ошибке
            }
        }

        /// <summary>
        /// Поиск по наименованию папки.
        /// </summary>
        static string GetFoldersSearchDirectory(string folder, string dirName)
        {
            try
            {
                return Directory.GetDirectories(folder, dirName, SearchOption.TopDirectoryOnly).FirstOrDefault();
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Ошибка при поиске папки {dirName} в {folder}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Получение всех файлов в папке.
        /// </summary>
        static string[] GetFileshDirectory(string folder)
        {
            try
            {
                return Directory.GetFiles(folder);
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Ошибка получении файлов в папке {folder}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Поиск файла по названию.
        /// </summary>
        static string GetFileSearchDirectory(string directory, string searchPattern)
        {
            try
            {
                return Directory.GetFiles(directory, searchPattern, SearchOption.TopDirectoryOnly).FirstOrDefault();
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Ошибка при поиске файлов в папке {directory}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Проверяет и создает указанные директории в baseFolder папке. Возвращает false при ошибке создания.
        /// </summary>
        static bool EnsureDirectoriesExist(string baseFolder, params string[] folderNames)
        {
            foreach (var folderName in folderNames)
            {
                string folderPath = Path.Combine(baseFolder, folderName);
                if (!CreateDirectoryWithLogging(folderPath))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Создает директорию по указанному пути, если она не существует, и логирует результат.
        /// </summary>
        static bool CreateDirectoryWithLogging(string path)
        {
            try
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                    Log(LogLevel.Debug, $"Папка {path} успешно создана.");
                }
                return true;
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Не удалось создать папку {path}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Создает папки для различных типов файлов и перемещает файлы в соответствующие папки.
        /// Если файл не является .xlsx, .pdf или .zip, он перемещается в папку "error".
        /// </summary>
        static FolderPaths CreateFolderMoveFiles(string creatingFolder, string[] arrayFiles)
        {
            // Папки для разных типов файлов
            var folderPaths = new FolderPaths
            {
                XlsxFolder = Path.Combine(creatingFolder, "xlsx"),
                PdfFolder = Path.Combine(creatingFolder, "pdf"),
                ZipFolder = Path.Combine(creatingFolder, "zip"),
                ErrorFolder = Path.Combine(creatingFolder, "error"),
                DocumentFolder = Path.Combine(creatingFolder, "document")
            };

            foreach (var file in arrayFiles)
            {
                try
                {
                    // Пропускаем папки (только файлы)
                    if (!File.Exists(file))
                    {
                        continue; // Это папка, пропускаем
                    }

                    string extension = Path.GetExtension(file).ToLower();
                    string destinationFolder = GetDestinationFolder(extension, folderPaths);
                    string destination = Path.Combine(destinationFolder, Path.GetFileName(file));

                    // Перемещаем файл
                    File.Move(file, destination);

                    // Логируем результат
                    if (extension == ".xlsx" || extension == ".pdf" || extension == ".zip")
                    {
                        Log(LogLevel.Debug, $"Перемещен файл {file} в {destinationFolder}");
                    }
                    else
                    {
                        Log(LogLevel.Warning, $"Файл {file} не является .xlsx, .pdf или .zip, перемещен в папку error.");
                    }
                }
                catch (Exception ex)
                {
                    Log(LogLevel.Error, $"Ошибка при перемещении файла {file}: {ex.Message}");
                }
            }
            return folderPaths;
        }

        /// <summary>
        /// Определяет папку назначения для файла в зависимости от его расширения.
        /// </summary>
        static string GetDestinationFolder(string extension, FolderPaths folderPaths)
        {
            string destinationFolder;

            switch (extension)
            {
                case ".xlsx":
                    destinationFolder = folderPaths.XlsxFolder;
                    break;
                case ".pdf":
                    destinationFolder = folderPaths.PdfFolder;
                    break;
                case ".zip":
                    destinationFolder = folderPaths.ZipFolder;
                    break;
                default:
                    destinationFolder = folderPaths.ErrorFolder;
                    break;
            }
            return destinationFolder;
        }

        /// <summary>
        /// Класс, представляющий пути к различным папкам для хранения файлов.
        /// </summary>
        public class FolderPaths
        {
            public string XlsxFolder { get; set; }
            public string PdfFolder { get; set; }
            public string ZipFolder { get; set; }
            public string ErrorFolder { get; set; }
            public string DocumentFolder { get; set; }
        }

        /// <summary>
        /// Метод парсинга Json файла заявки
        /// </summary>
        public static (string OrgTitle, string Title, string FormType, string ppudOrganization) ParseJsonFile(string filePath)
        {
            // Проверка существования файла
            if (!File.Exists(filePath))
            {
                Log(LogLevel.Fatal, $"Файл не найден: {filePath}");
                throw new FileNotFoundException($"Файл не найден: {filePath}");
            }

            Log(LogLevel.Debug, $"Начинается обработка файла: {filePath}");

            // Чтение содержимого файла
            string jsonContent;
            try
            {
                jsonContent = File.ReadAllText(filePath);
                Log(LogLevel.Debug, $"Файл успешно прочитан: {filePath}");
            }
            catch (Exception ex)
            {
                Log(LogLevel.Fatal, $"Ошибка чтения файла {filePath}: {ex.Message}");
                throw new IOException($"Ошибка чтения файла: {ex.Message}");
            }

            // Проверка на пустое содержимое
            if (string.IsNullOrWhiteSpace(jsonContent))
            {
                Log(LogLevel.Fatal, $"Файл пуст или содержит только пробелы: {filePath}");
                throw new InvalidOperationException("Файл пуст или содержит только пробелы.");
            }

            // Парсинг JSON
            try
            {
                JToken jsonToken = JToken.Parse(jsonContent);

                JObject jsonObject;
                if (jsonToken is JObject obj)
                {
                    jsonObject = obj;
                }
                else if (jsonToken is JArray array && array.Count > 0 && array[0] is JObject firstObj)
                {
                    jsonObject = firstObj; // Если JSON - массив, берем первый объект
                }
                else
                {
                    Log(LogLevel.Fatal, $"Неверный формат JSON: ожидался объект или массив объектов в файле {filePath}. Проверьте файл заявки.");
                    throw new InvalidOperationException("Неверный формат JSON: ожидался объект или массив объектов. Проверьте файл заявки.");
                }

                /// Извлечение значений с логированием ошибок
                string orgTitle = jsonObject?["orgFil"]?["title"]?.ToString();
                if (string.IsNullOrEmpty(orgTitle))
                {
                    Log(LogLevel.Fatal, $"Поле 'orgFil.title' отсутствует или пустое в JSON: {filePath}");
                    throw new InvalidOperationException("Поле 'orgFil.title' отсутствует или пустое.");
                }

                string title = jsonObject?["title"]?.ToString();
                if (string.IsNullOrEmpty(title))
                {
                    Log(LogLevel.Fatal, $"Поле 'title' отсутствует или пустое в JSON: {filePath}");
                    throw new InvalidOperationException("Поле 'title' отсутствует или пустое.");
                }

                string formType = jsonObject?["formTypeInt"]?["title"]?.ToString()?.Trim();
                if (string.IsNullOrEmpty(formType))
                {
                    Log(LogLevel.Fatal, $"Поле 'formTypeInt.title' отсутствует или пустое в JSON: {filePath}");
                    throw new InvalidOperationException("Поле 'formTypeInt.title' отсутствует или пустое.");
                }

                // Пытаемся найти организацию по названию
                var matchingKeyValue = _organizationValues.FirstOrDefault(kv => kv.Key == orgTitle);
                if (matchingKeyValue.Key == null)
                {
                    Log(LogLevel.Fatal, $"ППУД для организации [{orgTitle}] не найдена в коллекции _organizationValues. JSON: {filePath}");
                    throw new InvalidOperationException($"ППУД с ключом '{orgTitle}' не найдена.");
                }

                string ppudOrganization = matchingKeyValue.Value;

                return (orgTitle, title, formType, ppudOrganization);
            }
            catch (JsonReaderException ex)
            {
                Log(LogLevel.Error, $"Ошибка парсинга JSON в файле {filePath}: {ex.Message}");
                throw new InvalidOperationException($"Ошибка парсинга JSON: {ex.Message}");
            }
        }

        /// <summary>
        /// Метод проверяющий наличие файла xlsx и pdf, возвращает список, xlsx у которых нет pdf
        /// </summary>
        static string[] XlsxContainsPDF(string xlsxFolder, string pdfFolder)
        {
            // Получаем все файлы PDF в папке и создаем словарь по базовому имени
            var pdfFiles = Directory.GetFiles(pdfFolder, "*.pdf")
                                    .ToDictionary(pdfFile => Path.GetFileNameWithoutExtension(pdfFile), pdfFile => pdfFile);

            string[] xlsxFiles = Directory.GetFiles(xlsxFolder, "*.xlsx")
                                               .Where(file => !file.Contains("~$")) // Исключение временных файлов
                                               .ToArray();

            // Список для хранения путей к xlsx файлам, для которых нет соответствующего PDF
            List<string> xlsxWithoutPdf = new List<string>();

            // Перебираем файлы из папки xlsx
            foreach (var xlsxFile in xlsxFiles)
            {
                string xlsxName = Path.GetFileNameWithoutExtension(xlsxFile).Trim(); // Получаем имя файла без расширения

                // Удаляем слово "ОЦО" из имени файла, если оно есть
                string cleanedXlsxName = xlsxName.Replace("ОЦО", "").Trim();

                // Проверяем наличие суффикса "OK" или "ок" в конце строки
                bool hasOkSuffix = cleanedXlsxName.EndsWith("OK", StringComparison.OrdinalIgnoreCase) ||
                                   cleanedXlsxName.EndsWith("ОК", StringComparison.OrdinalIgnoreCase);

                // Убираем суффикс, если он есть в конце строки
                string baseName = hasOkSuffix
                    ? cleanedXlsxName.Substring(0, cleanedXlsxName.Length - 2) // Убираем последние два символа
                    : cleanedXlsxName;

                string normalizeName = baseName.ToLower().Replace(" ", string.Empty);

                // Проверка на наличие соответствующего PDF-файла
                bool hasMatchingPdf = pdfFiles.Any(pdfFile =>
                    pdfFile.Key.Trim().ToLower().Replace(" ", string.Empty).StartsWith(normalizeName));

                if (hasMatchingPdf)
                {
                    // Если в имени нет "OK", добавляем его в название
                    if (!hasOkSuffix)
                    {
                        string newXlsxName = $"{baseName} OK.xlsx";
                        string newXlsxPath = Path.Combine(xlsxFolder, newXlsxName);

                        // Переименовываем файл, добавляя "OK" в конец
                        File.Move(xlsxFile, newXlsxPath);

                        Log(LogLevel.Info, $"[*] Файл [{xlsxName}] переименован в [{newXlsxName}] и сохранен.");
                    }
                    Log(LogLevel.Debug, $"[+] Для файла [{xlsxName}] найден соответствующий PDF.");

                }
                else
                {
                    // Если PDF нет, добавляем файл в список
                    xlsxWithoutPdf.Add(xlsxFile);
                    Log(LogLevel.Warning, $"[-] Для файла [{xlsxName}] не найден соответствующий PDF. Файл добавлен в очередь на конвертирование.");
                }
            }

            // Возвращаем массив путей xlsx файлов, для которых нет PDF
            return xlsxWithoutPdf.ToArray();
        }

        /// <summary>
        /// Метод конвертации xlsx в pdf
        /// </summary>
        static void ConvertToPdf(IEnumerable<string> xlsxFiles, string outputFolder)
        {
            Excel.Application excelApplication = null;

            try
            {
                excelApplication = new Excel.Application();

                foreach (var file in xlsxFiles)
                {
                    Excel.Workbook workbook = null;

                    try
                    {
                        Log(LogLevel.Debug, $"Обработка файла: {file}");

                        // Проверка наличия суффикса OK или ОК
                        bool hasOkSuffix = file.EndsWith("OK", StringComparison.OrdinalIgnoreCase) ||
                                           file.EndsWith("ОК", StringComparison.OrdinalIgnoreCase);

                        // Формируем базовое имя без суффикса (если он есть)
                        string baseName = hasOkSuffix
                            ? file.Substring(0, file.Length - 2) // Убираем последние два символа
                            : file;

                        // Проверяем имя файла на наличие недопустимых символов
                        string sanitizedFileName = Path.GetFileNameWithoutExtension(baseName);
                        sanitizedFileName = string.Join("_", sanitizedFileName.Split(Path.GetInvalidFileNameChars()));

                        string outputFile = Path.Combine(outputFolder, $"{sanitizedFileName}.pdf");

                        Log(LogLevel.Debug, $"Выходной файл: {outputFile}");

                        // Открытие файла Excel
                        workbook = excelApplication.Workbooks.Open(file);

                        // Сохранение в PDF
                        workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputFile);

                        Log(LogLevel.Debug, $"Файл успешно конвертирован: {outputFile}");
                    }
                    catch (Exception ex)
                    {
                        Log(LogLevel.Error, $"Ошибка обработки файла '{file}': {ex.Message}");
                    }
                    finally
                    {
                        if (workbook != null)
                        {
                            workbook.Close(false);
                            Marshal.ReleaseComObject(workbook);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Ошибка при работе с Excel: {ex.Message}");
            }
            finally
            {
                if (excelApplication != null)
                {
                    excelApplication.Quit();
                    Marshal.ReleaseComObject(excelApplication);
                }

                // Принудительная сборка мусора для освобождения ресурсов
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// Метод нахождения и удаления процесса приложения
        /// </summary>
        static void KillExcelProcesses(string NameProceses)
        {
            try
            {
                string currentUser = Environment.UserName; // Получение имени текущего пользователя

                foreach (var process in Process.GetProcessesByName(NameProceses))
                {
                    try
                    {
                        if (IsProcessOwnedByCurrentUser(process))
                        {
                            Log(LogLevel.Debug, $"Завершаем процесс {NameProceses} с ID {process.Id}, пользователь: {currentUser}");
                            process.Kill();
                        }
                    }
                    catch (Exception ex)
                    {
                        Log(LogLevel.Error, $"Ошибка при завершении процесса {NameProceses} с ID {process.Id}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Ошибка при завершении процессов {NameProceses}: {ex.Message}");
            }
        }

        /// <summary>
        /// Метод нахождения процесса приложения по имени у текущей УЗ
        /// </summary>
        static bool IsProcessOwnedByCurrentUser(Process process)
        {
            try
            {
                // Проверка владельца процесса через WMI
                var query = $"SELECT * FROM Win32_Process WHERE ProcessId = {process.Id}";
                using (var searcher = new ManagementObjectSearcher(query))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        var outParams = obj.InvokeMethod("GetOwner", null, null);
                        if (outParams != null && outParams.Properties["User"] != null)
                        {
                            string user = outParams.Properties["User"].Value.ToString();
                            return string.Equals(user, Environment.UserName, StringComparison.OrdinalIgnoreCase);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Ошибка при определении владельца процесса {process.Id}: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// Метод перемещения профиля landocs
        /// </summary>
        static void MoveCustomProfileLandocs(string customFile, string landocsProfileFolder)
        {
            try
            {
                // Проверяем существование исходного файла
                if (!File.Exists(customFile))
                {
                    throw new FileNotFoundException($"Ошибка: исходный файл профиля landocs '{customFile}' не найден.");
                }

                // Убедимся, что папка назначения существует
                if (!Directory.Exists(landocsProfileFolder))
                {
                    throw new FileNotFoundException($"Ошибка: папка с профилями landocs '{customFile}' не найден.");
                }

                // Формируем полный путь к файлу в папке назначения
                string destinationFilePath = Path.Combine(landocsProfileFolder, Path.GetFileName(customFile));

                // Если файл назначения существует, меняем его расширение на .bak
                if (File.Exists(destinationFilePath))
                {
                    string backupFilePath = Path.ChangeExtension(destinationFilePath, ".bak");

                    // Удаляем старый .bak файл, если он существует
                    if (File.Exists(backupFilePath))
                    {
                        File.Delete(backupFilePath);
                    }

                    File.Move(destinationFilePath, backupFilePath);
                    Log(LogLevel.Debug, $"Выполнил резервную копию файла профиля [{destinationFilePath}] переименован в [{backupFilePath}].");
                }

                // Перемещаем новый файл
                File.Copy(customFile, destinationFilePath);

                Log(LogLevel.Debug, $"Кастомный файл профиля landocs успешно перемещен из '{customFile}' в '{destinationFilePath}'.");
            }
            catch (Exception ex)
            {
                // Логируем ошибку
                Log(LogLevel.Fatal, $"Ошибка перемещения профиля: {ex.Message}");

                // Бросаем исключение, чтобы завершить работу приложения
                throw new ApplicationException($"Критическая ошибка: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Метод запуска landocs
        /// </summary>
        public static IUIAutomationElement LaunchAndFindWindow(string appPath, string windowName, int maxWaitTime)
        {
            try
            {
                var automation = new CUIAutomation();
                var rootElement = automation.GetRootElement();

                Log(LogLevel.Info, $"Запуск приложения: {appPath}");
                var appProcess = Process.Start(appPath);

                if (appProcess == null)
                {
                    Log(LogLevel.Error, "Не удалось запустить приложение.");
                    throw new ApplicationException("Критическая ошибка: Не удалось запустить приложение.");
                }

                IUIAutomationElement appElement = null;
                int elapsedSeconds = 0;

                Log(LogLevel.Info, $"Поиск окна приложения с именем: [{windowName}]. Время ожидания:[{maxWaitTime}] сек.");

                while (elapsedSeconds < maxWaitTime && appElement == null)
                {
                    IUIAutomationCondition condition = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_NamePropertyId, windowName);
                    appElement = rootElement.FindFirst(TreeScope.TreeScope_Children, condition);

                    if (appElement == null)
                    {
                        Thread.Sleep(1000);
                        elapsedSeconds++;

                        Log(LogLevel.Debug, $"Ожидание окна приложения: [{windowName}]. Прошло [{elapsedSeconds}] секунд...");

                        // Каждые 10 секунд - лог уровня Info
                        if (elapsedSeconds % 10 == 0)
                        {
                            Log(LogLevel.Warning, $"Ожидание окна приложения: [{windowName}]. Прошло [{elapsedSeconds}] секунд.");
                        }
                    }
                }

                if (appElement != null)
                {
                    Log(LogLevel.Info, "Landocs успешно запустился.");
                }
                else
                {
                    Log(LogLevel.Error, "Окно приложения не найдено после максимального времени ожидания.");
                    throw new ApplicationException($"Критическая ошибка: Окно приложения '{windowName}' не найдено.");
                }

                return appElement;
            }
            catch (Exception ex)
            {
                Log(LogLevel.Fatal, $"Ошибка при запуске или поиске окна приложения: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Метод клика на элемент (Эмуляция программного нажатия)
        /// </summary>
        static void TryInvokeElement(IUIAutomationElement element)
        {
            try
            {
                if (element.GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId) is IUIAutomationInvokePattern invokePattern)
                {
                    if (element.CurrentIsEnabled != 0)
                    {
                        Task.Run(() =>
                        {
                            try
                            {
                                invokePattern.Invoke();
                                Console.WriteLine("Клик выполнен через Invoke.");
                            }
                            catch (COMException ex)
                            {
                                Console.WriteLine($"Ошибка COM во время Invoke: {ex.Message}");
                                
                            }
                        }).Wait(TimeSpan.FromSeconds(5)); // Устанавливаем тайм-аут
                    }
                    else
                    {
                        Console.WriteLine("Элемент недоступен для взаимодействия.");
                    }
                }
                else
                {
                    Console.WriteLine("Элемент не поддерживает InvokePattern.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Общая ошибка в TryInvokeElement: {ex.Message}");
            }
        }

        /// <summary>
        /// Метод клика на элемент (Эмуляция физического нажатия)
        /// </summary>
        static void ClickElementWithMouse(IUIAutomationElement element)
        {
            try
            {
                // Получение границ элемента
                object boundingRectValue = element.GetCurrentPropertyValue(UIA_PropertyIds.UIA_BoundingRectanglePropertyId);

                // Проверяем, что значение границ корректное
                if (!(boundingRectValue is double[] boundingRectangle) || boundingRectangle.Length != 4)
                {
                    Log(LogLevel.Warning, "Не удалось получить или обработать границы элемента.");
                    throw new InvalidOperationException("Некорректные границы элемента.");
                }

                // Извлечение координат
                int left = (int)boundingRectangle[0];
                int top = (int)boundingRectangle[1];
                int right = (int)boundingRectangle[2];
                int bottom = (int)boundingRectangle[3];

                // Проверяем, что размеры валидны
                /*if (right <= left || bottom <= top)
                {
                    Log(LogLevel.Warning, "Границы элемента некорректны.");
                    throw new InvalidOperationException("Неверные размеры элемента.");
                }*/

                // Расчет центра элемента
                int x = left + right / 2;
                int y = top + bottom / 2;

                // Устанавливаем курсор на центр элемента
                if (!SetCursorPos(x, y))
                {
                    Log(LogLevel.Error, $"Не удалось установить курсор на позицию: X={x}, Y={y}");
                    throw new InvalidOperationException("Ошибка установки позиции курсора.");
                }

                // Небольшая задержка перед кликом
                Thread.Sleep(100);

                // Выполняем клик
                mouse_event((int)MouseFlags.LeftDown, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(200);
                mouse_event((int)MouseFlags.LeftUp, 0, 0, 0, UIntPtr.Zero);

                Log(LogLevel.Info, $"Клик выполнен по элементу в центре: X={x}, Y={y}");
            }
            catch (COMException ex)
            {
                Log(LogLevel.Error, $"COM-ошибка при попытке кликнуть по элементу: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Log(LogLevel.Error, $"Общая ошибка при клике по элементу: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Метод поиска элемента по xpath
        /// </summary>
        static IUIAutomationElement FindElementByXPath(IUIAutomationElement root, string xpath, int secondsToWait)
        {
            var automation = new CUIAutomation();
            IUIAutomationCondition trueCondition = automation.CreateTrueCondition();
            string[] parts = xpath.Split('/');
            IUIAutomationElement currentElement = root;

            int elapsedSeconds = 0;
            const int checkInterval = 500;

            while (elapsedSeconds < secondsToWait)
            {
                foreach (var part in parts)
                {
                    if (currentElement == null)
                    {
                        Console.WriteLine("Текущий элемент равен null, поиск прерван.");
                        return null;
                    }

                    var split = part.Split(new char[] { '[', ']' }, StringSplitOptions.RemoveEmptyEntries);
                    string type = split[0];
                    int index = split.Length > 1 ? int.Parse(split[1]) - 1 : 0;

                    // Проверяем, что мы можем найти дочерние элементы
                    IUIAutomationElementArray children = currentElement.FindAll(TreeScope.TreeScope_Children, trueCondition);

                    if (children == null || children.Length == 0)
                    {
                        Console.WriteLine("Дочерние элементы не найдены.");
                        return null;
                    }

                    bool found = false;
                    int typeCount = 0;

                    for (int i = 0; i < children.Length; i++)
                    {
                        IUIAutomationElement child = children.GetElement(i);

                        if (child != null && child.CurrentControlType == GetControlType(type))
                        {
                            if (typeCount == index)
                            {
                                currentElement = child;
                                found = true;
                                break;
                            }
                            typeCount++;
                        }
                    }

                    if (!found)
                    {
                        currentElement = null;
                        break;
                    }
                }

                if (currentElement != null)
                {
                    return currentElement;
                }

                Thread.Sleep(checkInterval);
                elapsedSeconds += checkInterval / 1000;
            }

            return null;
        }

        /// <summary>
        /// Метод поиска элемента по параметру Name
        /// </summary>
        static IUIAutomationElement FindElementByName(IUIAutomationElement root, string name, int secondsToWait)
        {
            var automation = new CUIAutomation();
            IUIAutomationCondition nameCondition = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_NamePropertyId, name);
            int elapsedSeconds = 0;
            const int checkInterval = 500;

            while (elapsedSeconds < secondsToWait)
            {
                IUIAutomationElement element = root.FindFirst(TreeScope.TreeScope_Descendants, nameCondition);

                if (element != null)
                {
                    return element;
                }

                Thread.Sleep(checkInterval);
                elapsedSeconds += checkInterval / 1000;
            }

            return null;
        }

        /// <summary>
        /// Метод возвращающий тип ControlType
        /// </summary>
        static int GetControlType(string type)
        {
            type = type.ToLower();

            switch (type)
            {
                case "pane": return UIA_ControlTypeIds.UIA_PaneControlTypeId;
                case "table": return UIA_ControlTypeIds.UIA_TableControlTypeId;
                case "tab": return UIA_ControlTypeIds.UIA_TabControlTypeId;
                case "tabitem": return UIA_ControlTypeIds.UIA_TabItemControlTypeId;
                case "button": return UIA_ControlTypeIds.UIA_ButtonControlTypeId;
                case "group": return UIA_ControlTypeIds.UIA_GroupControlTypeId;
                case "checkbox": return UIA_ControlTypeIds.UIA_CheckBoxControlTypeId;
                case "combobox": return UIA_ControlTypeIds.UIA_ComboBoxControlTypeId;
                case "edit": return UIA_ControlTypeIds.UIA_EditControlTypeId;
                case "text": return UIA_ControlTypeIds.UIA_TextControlTypeId;
                case "window": return UIA_ControlTypeIds.UIA_WindowControlTypeId;
                case "custom": return UIA_ControlTypeIds.UIA_CustomControlTypeId;
                case "tree": return UIA_ControlTypeIds.UIA_TreeControlTypeId;
                case "toolbar": return UIA_ControlTypeIds.UIA_ToolBarControlTypeId;
                case "dataitem": return UIA_ControlTypeIds.UIA_DataItemControlTypeId;
                default: return UIA_ControlTypeIds.UIA_PaneControlTypeId;
            }
        }

        /// <summary>
        /// Метод возвращающий элемент на котором сейчас установлен фокус
        /// </summary>
        static IUIAutomationElement GetFocusedElement()
        {
            var automation = new CUIAutomation();
            IUIAutomationElement focusedElement = automation.GetFocusedElement();

            if (focusedElement != null)
            {
                try
                {
                    Console.WriteLine("Элемент с фокусом найден:");
                    Console.WriteLine($"Имя элемента: {focusedElement.CurrentName}");
                    Console.WriteLine($"Тип элемента: {focusedElement.CurrentControlType}");
                    Console.WriteLine($"Тип элемента: {focusedElement.CurrentLocalizedControlType}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка при получении информации об элементе с фокусом: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Элемент с фокусом не найден.");
            }

            return focusedElement;
        }

        /// <summary>
        /// Метод возвращающий элемент окна с ошибкой
        /// </summary>
        static IUIAutomationElement GetErrorWindowElement(IUIAutomationElement rootElement, string echildrenNameWindow)
        {
            var targetWindowError = FindElementByName(rootElement, echildrenNameWindow, 60);

            // Проверяем значение свойства Name элемента
            if (targetWindowError != null)
            {
                // Создаем условия для поиска title и message
                var automation = new CUIAutomation();

                // Условие для поиска элемента сообщения (message)
                var messageCondition = automation.CreatePropertyCondition(
                    UIA_PropertyIds.UIA_ControlTypePropertyId,
                    UIA_ControlTypeIds.UIA_TextControlTypeId
                );
                var messageElement = targetWindowError.FindFirst(TreeScope.TreeScope_Children, messageCondition);

                string message = messageElement != null
                    ? messageElement.GetCurrentPropertyValue(UIA_PropertyIds.UIA_NamePropertyId) as string
                    : "Сообщение не найдено";
                Log(LogLevel.Fatal, $"Появилось окно [Ошибка], текст сообщения: [{message}]");
                // Ищем кнопку "ОК"
                var buttonOk = FindElementByName(targetWindowError, "&ОК", 60);

                throw new Exception("Появилось окно ошибки. Работа робота завершена.");
            }
            else
            {
                throw new Exception($"Появилось окно ошибки. Не удалось определить элемент. Робот завершает работу.");
            }
        }

        /// <summary>
        /// Метод возвращающий ключ контрагента найденного по ИНН и КПП
        /// </summary>
        static int? FindCounterpartyKey(Dictionary<int, string[]> counterpartyElements, string innValue, string kppValue, string counterpartyName = null)
        {
            // Приводим значения ИНН и КПП к единому формату заранее
            string formattedInnValue = $"ИНН:{innValue}".Replace(" ", "").Trim().ToLower();
            string formattedKppValue = string.IsNullOrEmpty(kppValue) ? null : $"КПП:{kppValue}".Replace(" ", "").Trim().ToLower();
            string formattedCounterpartyName = string.IsNullOrEmpty(counterpartyName) ? null : counterpartyName.Replace(" ", "").Trim().ToLower();

            foreach (var kvp in counterpartyElements)
            {
                // Очищаем элементы списка контрагентов от лишних пробелов и приводим к нижнему регистру один раз
                var formattedElements = kvp.Value.Select(x => x.Replace(" ", "").Trim().ToLower()).ToList();

                // Проверяем наличие ИНН
                bool innMatch = formattedElements.Contains(formattedInnValue);

                // Проверяем наличие КПП (если оно задано)
                bool kppMatch = string.IsNullOrEmpty(formattedKppValue) || formattedElements.Contains(formattedKppValue);

                // Если КПП отсутствует, проверяем по имени контрагента
                bool nameMatch = string.IsNullOrEmpty(formattedKppValue) && !string.IsNullOrEmpty(formattedCounterpartyName) &&
                                 formattedElements.Any(x => x.Contains(formattedCounterpartyName));

                // Если найдено совпадение по ИНН и либо КПП, либо имени
                if (innMatch && (kppMatch || nameMatch))
                {
                    return kvp.Key;
                }
            }
            return null; // Возвращаем null, если совпадений не найдено
        }

        /// <summary>
            /// Метод возвращающий параметры с названия файла для landocs
            /// </summary>
        static FileData GetParseNameFile(string fileName)
        {
            // Регулярное выражение для парсинга строки
            var match = Regex.Match(fileName,
                @"Акт св П \d+\s+(.*?)\s+№(\S+)\s+(\d{2}\.\d{2}\.\d{2})_(\d+)_?(\d+)?");

            if (match.Success)
            {
                return new FileData
                {
                    CounterpartyName = match.Groups[1].Value.Trim(),
                    Number = match.Groups[2].Value.Trim(),
                    FileDate = match.Groups[3].Value.Trim(),
                    INN = match.Groups[4].Value.Trim(),
                    KPP = match.Groups[5].Success ? match.Groups[5].Value.Trim() : null
                };
            }
            else
            {
                Console.WriteLine($"Не удалось распознать файл: {fileName}");
                //Добавить перемещение в папку error
                return null;
            }
        }

        /// <summary>
        /// Класс, с параметрами файла для landocs
        /// </summary>
        public class FileData
        {
            public string CounterpartyName { get; set; }
            public string Number { get; set; }
            public string FileDate { get; set; }
            public string INN { get; set; }
            public string KPP { get; set; }
        }
        #endregion
    }
}