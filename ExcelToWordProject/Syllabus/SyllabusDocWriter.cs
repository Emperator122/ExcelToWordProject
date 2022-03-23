using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelToWordProject.Models;
using ExcelToWordProject.Syllabus.Tags;
using ExcelToWordProject.Utils;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace ExcelToWordProject.Syllabus
{
    internal class SyllabusDocWriter : IDisposable
    {
        // Список смарт тегов, для которых будут пробегаться параграфы
        private readonly SmartTagType[] _smartTagParagraphJob = {SmartTagType.Content, SmartTagType.ExtendedContent};


        // Список тегов в таблицах текущего шаблона
        private Dictionary<int, List<BaseSyllabusTag>> _tablesTags;

        // Список TextBlock тегов в таблицах текущего шаблона
        private Dictionary<int, List<TextBlockTag>> _tablesTextBlockTags;


        public SyllabusDocWriter(SyllabusExcelReader syllabusExcelReader, SyllabusParameters parameters)
        {
            SyllabusExcelReader = syllabusExcelReader;
            Parameters = parameters;
        }

        public SyllabusParameters Parameters { get; set; }
        public SyllabusExcelReader SyllabusExcelReader { get; set; }

        public void Dispose()
        {
            SyllabusExcelReader.Dispose();
        }

        /// <summary>
        ///     Конвертация Excel документа в Word со всеми тегами
        /// </summary>
        /// <param name="resultFolderPath">Путь к папке, куда будут сложены результаты</param>
        /// <param name="baseDocumentPath">Пусть к шаблону</param>
        /// <param name="fileNamePrefix">Префикс к имени файла</param>
        /// <param name="progress">Прогресс (для отображения вне потока)</param>
        public void ConvertToDocx(string resultFolderPath, string baseDocumentPath, string fileNamePrefix = "",
            IProgress<int> progress = null)
        {
            var baseDocument = DocX.Load(baseDocumentPath);

            FindTablesTags(baseDocument);

            if (Parameters.Tags.HasActiveSmartTags())
                SmartTagsDocumentHandler(resultFolderPath, baseDocumentPath, fileNamePrefix, progress);
            else // просто заменяем теги
                DefaultTagsDocumentHandler(resultFolderPath, baseDocument, fileNamePrefix, progress);
        }

        protected void SmartTagsDocumentHandler(string resultFolderPath, string baseDocumentPath,
            string fileNamePrefix = "", IProgress<int> progress = null)
        {
            // Есть ли смарт-теги, работающие со списком компетенций
            var hasSmartModulesContentTags =
                Parameters.Tags.FindIndex(el =>
                    el is SmartSyllabusTag && el.ListName == Parameters.ModulesContentListName) != -1;

            // получаем все модули
            var modules = SyllabusExcelReader.GetAllModules(Parameters.ModulesYears);
            var i = 0;
            foreach (var module in modules)
            {
                // Репортим прогресс
                i++;
                progress?.Report(i * 100 / modules.Count);

                // Приготовим имя и путь к файлу
                var safeName =
                    PathUtils.RemoveIllegalFileNameCharacters(fileNamePrefix + module.Index + " " + module.Name +
                                                              ".docx");
                safeName = PathUtils.FixFileNameLimit(safeName);
                var resultFilePath = Path.Combine(resultFolderPath, safeName);

                // Создадим новый файл с результатом
                var doc = PathUtils.CopyFile(baseDocumentPath, resultFilePath);
                if (doc == null)
                    continue;

                // Обработаем данный модуль
                ModuleHandler(doc, module, hasSmartModulesContentTags);

                // Сохраняем файл
                doc.Save();
                doc.Dispose();
            }
        }

        protected void DefaultTagsDocumentHandler(string resultFolderPath, DocX baseDocument,
            string fileNamePrefix = "", IProgress<int> progress = null)
        {
            fileNamePrefix = fileNamePrefix == "" ? "UnsetFileName" : fileNamePrefix;
            var safeName = PathUtils.RemoveIllegalFileNameCharacters(fileNamePrefix + ".docx");
            var resultFilePath = Path.Combine(resultFolderPath, safeName);
            baseDocument.SaveAs(resultFilePath);
            var doc = DocX.Load(resultFilePath);

            TablesHandler(doc); // TODO: Зачем?

            TextBlocksHandler(doc);

            // бежим по списку тегов
            foreach (var tag in Parameters.Tags.Where(tag => tag is DefaultSyllabusTag && tag.Active))
                doc.ReplaceText(tag.Tag, tag.GetValue(excelData: SyllabusExcelReader.ExcelData));
            doc.Save();
            doc.Dispose();
            progress?.Report(100);
        }

        /// <summary>
        ///     Запись информации об одной дисциплине в docx
        /// </summary>
        /// <param name="doc">Документ</param>
        /// <param name="module">Дисциплина</param>
        /// <param name="hasSmartModulesContentTags">Есть ли активные теги, работающие с контентом</param>
        protected void ModuleHandler(DocX doc, Module module, bool hasSmartModulesContentTags)
        {
            // получаем компентенции
            List<Content> contentList = null;
            if (hasSmartModulesContentTags)
                contentList = SyllabusExcelReader.ParseContentList(module);

            // обработка таблиц
            TablesHandler(doc, module, contentList);

            // обработка TextBlock
            TextBlocksHandler(doc, module, contentList);

            // бежим по списку тегов
            foreach (var tag in Parameters.Tags)
            {
                // и заполняем каждый активный тег
                if (!tag.Active) continue;
                // получаем значение тега
                var tagValue = tag.GetValue(module, contentList, SyllabusExcelReader.ExcelData);

                if (tag is SmartSyllabusTag smartTag) // если у нас контент-тег, то заполняем параграфы
                {
                    // Для некоторых тегов используется доп. хендлер
                    if (_smartTagParagraphJob.Contains(smartTag.Type))
                        ArrayToParagraphs(doc, smartTag, tagValue.Split('\n'));
                }

                doc.ReplaceText(tag.Tag, tagValue);
            }
        }

        /// <summary>
        ///     Поиск всех тегов, находящихся в последних строках таблиц.
        ///     В результате будут заполнены поля класса tablesTags и tablesTextBlockTags,
        ///     представляющие из себя листы листов тегов.
        ///     Листов по количество таблиц в документе,
        ///     внутренних листов по количеству тегов в каждой таблице.
        /// </summary>
        /// <param name="doc">Документ</param>
        /// <param name="refresh">Следует ли обновить поля</param>
        protected void FindTablesTags(DocX doc, bool refresh = true)
        {
            if (_tablesTags != null && _tablesTextBlockTags != null && !refresh)
                return;

            _tablesTags = new Dictionary<int, List<BaseSyllabusTag>>();
            for (var i = 0; i < doc.Tables.Count; i++)
            {
                var table = doc.Tables[i];

                var row = table.Rows.Last();

                // найдем все Smart и Default теги в последней ровке траблицы
                var tags =
                    Parameters
                        .Tags
                        .FindAll(
                            tag =>
                                tag.Active &&
                                row.FindAll(tag.Tag).Count > 0
                        );

                if (tags.Count > 0)
                    _tablesTags[i] = tags;

                // найдем TextBlock теги
                _tablesTextBlockTags = new Dictionary<int, List<TextBlockTag>>();
                var textBlockTagsFilteredGroup =
                    Parameters
                        .TextBlockTags
                        .OrderedByPriority()
                        .GroupedByKey()
                        .Where(
                            group =>
                                (group.FirstOrDefault()?.Active ?? false) &&
                                row.FindAll(group.First().Tag).Count > 0
                        );
                var textBlockTagsMixedList = new List<TextBlockTag>();
                foreach (var textBlockTagGroup in textBlockTagsFilteredGroup)
                    textBlockTagsMixedList.AddRange(textBlockTagGroup);

                if (textBlockTagsMixedList.Count > 0)
                    _tablesTextBlockTags[i] = textBlockTagsMixedList;
            }
        }

        /// <summary>
        ///     Обработка таблиц в документе.
        ///     Достает необходимые значения из тегов, чтобы заполнить
        ///     очередную таблицу, а затем заполнить ее в другом методе.
        /// </summary>
        /// <param name="doc">Документ</param>
        /// <param name="module">Информация о дисциплине</param>
        /// <param name="contentList">Компетенции дисциплины</param>
        protected void TablesHandler(DocX doc, Module module = null, List<Content> contentList = null)
        {
            // Находим все теги, содержащиеся в таблицах
            FindTablesTags(doc, false);
            foreach (var kv in _tablesTags)
            {
                var table = doc.Tables[kv.Key];
                var tags = kv.Value;

                var textBlockTags = new List<TextBlockTag>();
                if (_tablesTextBlockTags.ContainsKey(kv.Key))
                    textBlockTags.AddRange(_tablesTextBlockTags[kv.Key]);

                // Удалим TextBlock теги не удовлетворяющие условиям
                for (var i = 0; i < textBlockTags.Count; i++)
                {
                    var isValid = textBlockTags[i]
                        .CheckConditions(
                            Parameters.Tags,
                            module,
                            contentList,
                            SyllabusExcelReader.ExcelData
                        );
                    if (isValid) continue;
                    textBlockTags.RemoveAt(i);
                    i--;
                }

                // Сгруппируем TextBlock теги по ключу
                var textBlockTagsGroup = textBlockTags.OrderedByPriority().GroupedByKey().ToList();

                // Массив значений тегов, оставим доп место для TextBlock тегов
                var tagsValues = new string[tags.Count + textBlockTagsGroup.Count][];

                // Обработка Smart и Default тегов
                for (var i = 0; i < tags.Count; i++)
                    tagsValues[i] =
                        tags[i]
                            .GetValue(module, contentList, SyllabusExcelReader.ExcelData).Split('\n');

                // Обработка TextBlock тегов
                var maxLength = tagsValues.Max(arr => arr?.Length ?? 0);
                var k = tags.Count;
                foreach (var group in textBlockTagsGroup) // Для каждой группы ключей TextBlock тегов
                {
                    // Массив с результатом, равный максимальной
                    // высоте столбца таблицы
                    var values = new string[maxLength];
                    var defaultValue = group.FirstOrDefault()?.DefaultValue ?? "";
                    for (var i = 0; i < values.Length; i++)
                        values[i] = defaultValue;

                    // Поставим значения TextBlock тегов в нужных строках таблицы
                    foreach (var textBlockTag in group)
                    {
                        if (textBlockTag.IsDefault) continue;
                        // i,j - бежим по всей таблице
                        for (var i = 0; i < maxLength; i++)
                        {
                            var allConditionsSatisfied = true;
                            for (var j = 0; j < tags.Count; j++)
                            {
                                // если TextBlock не зависит от текущего Base тега, то пропускаем итерацию
                                var condition =
                                    textBlockTag
                                        .Conditions
                                        .FirstOrDefault(
                                            textBlockCondition => textBlockCondition.TagName == tags[j].Key
                                        );
                                if (condition == null)
                                    continue;

                                // иначе получаем текущее значение Base тега в таблице
                                string rowTagValue = null;
                                if (i < tagsValues[j]?.Length)
                                    rowTagValue = tagsValues[j][i];
                                else if (i >= tagsValues[j]?.Length)
                                    rowTagValue = tagsValues[j][tagsValues[j].Length - 1];

                                // и сравниваем значение Base тега с условием TextBlock тега
                                if (condition.Condition == rowTagValue) continue;
                                allConditionsSatisfied = false;
                                break;
                            }

                            // Если не зафакапился флаг, то тег становится на свое место
                            // (по идее)
                            if (!allConditionsSatisfied) continue;
                            values[i] = textBlockTag.GetValue2();
                            break;
                        }
                    }

                    tagsValues[k] = values;
                    k++;
                }

                // В итоге смешиваем Base теги с TextBlock тегами (первыми из них)
                var mixedTags = new List<BaseSyllabusTag>(tags);
                mixedTags.AddRange(
                    textBlockTagsGroup
                        .Select(group => group.First())
                );

                // Заполняем таблицу
                FillTable(table, mixedTags, tagsValues);
            }
        }

        /// <summary>
        ///     Заполнение конкретной таблицы.
        /// </summary>
        /// <param name="table">Таблица</param>
        /// <param name="tags">Теги</param>
        /// <param name="tagsValues">Значения тегов</param>
        protected void FillTable(Table table, List<BaseSyllabusTag> tags, string[][] tagsValues)
        {
            try
            {
                var patternRowIndex = table.Rows.Count - 1;
                var patternRow = table.Rows.Last();

                var maxLength = tagsValues.Max(arr => arr?.Length ?? 0);

                for (var i = 0; i < maxLength; i++)
                {
                    table.InsertRow(patternRow);
                    var currentRow = table.Rows.Last();

                    for (var j = 0; j < tagsValues.Length; j++)
                        if (i < tagsValues[j]?.Length)
                            currentRow.ReplaceText(tags[j].Tag, tagsValues[j][i]);
                        else if (i >= tagsValues[j]?.Length)
                            currentRow.ReplaceText(tags[j].Tag, tagsValues[j][tagsValues[j].Length - 1]);
                }

                table.RemoveRow(patternRowIndex);
            }
            catch
            {
                Console.WriteLine(@"Ошибка при обработке таблицы");
            }
        }

        /// <summary>
        ///     Реализация костыля для раздела компетенций.
        ///     Заменяет все одинокие теги на параграфы из tagValue, сплитнутого по \n.
        /// </summary>
        /// <param name="doc">Документ DocX</param>
        /// <param name="contentTag">Тег</param>
        /// <param name="lines">Массив строк, которые станут параграфами на месте тега</param>
        protected void ArrayToParagraphs(DocX doc, SmartSyllabusTag contentTag, string[] lines)
        {
            try
            {
                // расставим знаки препинания
                if (lines.Length > 0)
                {
                    for (var i = 0; i < lines.Length - 1; i++)
                        lines[i] += ";";
                    lines[lines.Length - 1] += ".";
                }

                // добавление параграфов
                var tagString = contentTag.Tag;
                foreach (var paragraph in doc.Paragraphs)
                    if (paragraph.Text.Trim() == tagString)
                    {
                        var p = paragraph.InsertParagraphAfterSelf(paragraph);
                        for (var i = 0; i < lines.Length; i++)
                        {
                            p.ReplaceText(tagString, lines[i]);
                            if (i != lines.Length - 1)
                                p = p.InsertParagraphAfterSelf(paragraph);
                        }

                        paragraph.Remove(false);
                    }
            }
            catch
            {
                Console.WriteLine(@"Ошибка при заполнении параграфа");
            }
        }

        protected void TextBlocksHandler(DocX doc, Module module = null, List<Content> contentList = null)
        {
            var textBlockTagsGroups = Parameters.TextBlockTags.OrderedByPriority().GroupedByKey();

            foreach (var textBlockTagsGroup in textBlockTagsGroups)
            {
                if (!textBlockTagsGroup.FirstOrDefault()?.Active ?? true) continue; // если нет активных тегов в группе
                var isValid = false;
                TextBlockTag outTextBlockTag = null;
                foreach (var textBlockTag in textBlockTagsGroup)
                {
                    outTextBlockTag = textBlockTag;
                    if(textBlockTag.IsDefault) continue;
                    isValid = textBlockTag.CheckConditions(Parameters.Tags, module, contentList,
                        SyllabusExcelReader.ExcelData);
                    if (isValid) break;
                }

                if (outTextBlockTag == null)
                    continue;

                var tagValue = isValid ? outTextBlockTag.GetValue2() : (outTextBlockTag.DefaultValue ?? "");
                doc.ReplaceText(outTextBlockTag.Tag, tagValue);
            }
        }
    }
}