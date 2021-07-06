using ExcelDataReader;
using ExcelToWordProject.Models;
using ExcelToWordProject.Syllabus;
using ExcelToWordProject.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace ExcelToWordProject.Syllabus
{
    class SyllabusDocWriter : IDisposable
    {
        public SyllabusParameters Parameters { get; set;  }
        public SyllabusExcelReader SyllabusExcelReader { get; set; }


        // Список тегов в таблицах текущего шаблона
        Dictionary<int, List<BaseSyllabusTag>> tablesTags;

        // Список смарт тегов, для которых будут пробегаться параграфы
        readonly SmartTagType[] smartTagParagraphJob = { SmartTagType.Content, SmartTagType.ExtendedContent };


        public SyllabusDocWriter(SyllabusExcelReader syllabusExcelReader, SyllabusParameters parameters)
        {
            SyllabusExcelReader = syllabusExcelReader;
            Parameters = parameters;           
        }

        public void Dispose()
        {
            SyllabusExcelReader.Dispose();
        }

        /// <summary>
        /// Конвертация Excel документа в Word со всеми тегами
        /// </summary>
        /// <param name="resultFolderPath">Путь к папке, куда будут сложены результаты</param>
        /// <param name="baseDocumentPath">Пусть к шаблону</param>
        /// <param name="fileNamePrefix">Префикс к имени файла</param>
        /// <param name="progress">Прогресс (для отображения вне потока)</param>
        public void ConvertToDocx(string resultFolderPath, string baseDocumentPath, string fileNamePrefix = "", IProgress<int> progress = null)
        {
            DocX baseDocument = DocX.Load(baseDocumentPath);

            FindTablesTags(baseDocument);

            if (Parameters.HasActiveSmartTags)
            {
                // Есть ли смарт-теги, работающие со списком компетенций
                bool hasSmartModulesContentTags =
                    Parameters.Tags.FindIndex(el => 
                    el is SmartSyllabusTag && el.ListName == Parameters.ModulesContentListName) != -1;

                // получаем все модули
                List<Module> modules = SyllabusExcelReader.GetAllModules(Parameters.ModulesYears);
                int i = 0;
                foreach(Module module in modules) 
                {
                    // Репортим прогресс
                    i++;
                    if (progress != null)
                        progress.Report(i * 100 / modules.Count());

                    // Приготовим имя и путь к файлу
                    string safeName = PathUtils.RemoveIllegalFileNameCharacters(fileNamePrefix + module.Index + " " + module.Name + ".docx");
                    safeName = PathUtils.FixFileNameLimit(safeName);
                    string resultFilePath = Path.Combine(resultFolderPath, safeName);

                    // Создадим новый файл с результатом
                    DocX doc = PathUtils.CopyFile(baseDocumentPath, resultFilePath);
                    if (doc == null)
                        continue;

                    // Обработаем данный модуль
                    ModuleHandler(doc, module, hasSmartModulesContentTags);
                   
                    // Сохраняем файл
                    doc.Save();
                    doc.Dispose();
                }
            }
            else // просто заменяем теги
            {
                fileNamePrefix = fileNamePrefix == "" ? "UnsetFileName" : fileNamePrefix;
                string safeName = PathUtils.RemoveIllegalFileNameCharacters(fileNamePrefix +".docx");
                string resultFilePath = Path.Combine(resultFolderPath, safeName);
                baseDocument.SaveAs(resultFilePath);
                DocX doc = DocX.Load(resultFilePath);

                TablesHandler(doc);
                // бежим по списку тегов
                foreach (BaseSyllabusTag tag in Parameters.Tags)
                {
                    // и заполняем каждый активный тег
                    if (tag is DefaultSyllabusTag && tag.Active)
                        doc.ReplaceText(tag.Tag, tag.GetValue(excelData:SyllabusExcelReader.ExcelData));

                }
                doc.Save();
                doc.Dispose();
                if (progress != null)
                    progress.Report(100);
            }
        }

        /// <summary>
        /// Запись информации об одной дисциплине в docx
        /// </summary>
        /// <param name="doc">Документ</param>
        /// <param name="module">Дисциплина</param>
        /// <param name="hasSmartModulesContentTags">Есть ли активные теги, работающие с контентом</param>
        /// <param name="hasSmartPlanListTags">Есть ли активные теги, работающие с листом "План"</param>
        protected void ModuleHandler(DocX doc, Module module, bool hasSmartModulesContentTags)
        {
            // получаем компентенции
            List<Content> contentList = null;
            if (hasSmartModulesContentTags)
                contentList = SyllabusExcelReader.ParseContentList(module);

            TablesHandler(doc, module, contentList);

            // бежим по списку тегов
            foreach (var tag in Parameters.Tags)
            {
                // и заполняем каждый активный тег
                string tagValue = "err";
                if (tag.Active)
                {
                    // получаем значение тега
                    tagValue = tag.GetValue(module, contentList, SyllabusExcelReader.ExcelData);

                    if (tag is SmartSyllabusTag) // если у нас контент-тег, то заполняем параграфы
                    {
                        SmartSyllabusTag smartTag = tag as SmartSyllabusTag;
                        // Для некоторых тегов используется доп. хендлер
                        if (smartTagParagraphJob.Contains(smartTag.Type))
                            ArrayToParagraps(doc, smartTag, tagValue.Split('\n'));
                    }

                    doc.ReplaceText(tag.Tag, tagValue);
                }
            }
        }
        
        /// <summary>
        /// Поиск всех тегов, находящихся в последних строках таблиц.
        /// В результате будет заполнено поле класса tablesTags,
        /// представляющее из себя лист листов тегов.
        /// 
        /// Листов по количество таблиц в документе,
        /// внутренних листов по количеству тегов в каждой таблице.
        /// </summary>
        /// <param name="doc">Документ</param>
        /// <param name="refresh">Следует ли обновить поле</param>
        protected void FindTablesTags(DocX doc, bool refresh = true)
        {
            if (tablesTags != null && !refresh)
                return;

            tablesTags = new Dictionary<int, List<BaseSyllabusTag>>();
            for (int i = 0; i < doc.Tables.Count; i++)
            {
                var table = doc.Tables[i];
                Row row = table.Rows.Last();
                List<BaseSyllabusTag> tags = Parameters.Tags.FindAll(tag
                    => tag.Active && row.FindAll(tag.Tag).Count > 0);
                if(tags.Count > 0)
                    tablesTags[i] = tags;
            }
        }
        
        /// <summary>
        /// Обработка таблиц в документе.
        /// Достает необходимые значения из тегов, чтобы заполнить
        /// очередную таблицу, а затем заполнить ее в другом методе.
        /// </summary>
        /// <param name="doc">Документ</param>
        /// <param name="module">Информация о дисциплине</param>
        /// <param name="contentList">Компетенции дисциплины</param>
        /// <param name="properties">Другие свойства дисциплины</param>
        protected void TablesHandler(DocX doc, Module module=null, List<Content> contentList=null)
        {
            // Находим все теги, содержащиеся в таблицах
            FindTablesTags(doc, false);
            foreach (var kv in tablesTags)
            {
                var table = doc.Tables[kv.Key];
                List<BaseSyllabusTag> tags = kv.Value;

                // Заполняем массив значений тегов
                string[][] tagsValues = new string[tags.Count][];
                for (int i = 0; i < tags.Count; i++)
                    tagsValues[i] = tags[i].GetValue(module, contentList, SyllabusExcelReader.ExcelData).Split('\n');

                // Заполняем таблицу
                FillTable(table, tags, tagsValues);
            }

        }

        /// <summary>
        /// Заполнение конкретной таблицы.
        /// </summary>
        /// <param name="table">Таблица</param>
        /// <param name="tags">Теги</param>
        /// <param name="tagsValues">Значения тегов</param>
        protected void FillTable(Table table, List<BaseSyllabusTag> tags, string[][] tagsValues)
        {

            try
            {
                int patternRowIndex = table.Rows.Count() - 1;
                Row patternRow = table.Rows.Last();

                int maxLength = tagsValues.Max((arr) => arr?.Length ?? 0);

                for (int i = 0; i < maxLength; i++)
                {
                    table.InsertRow(patternRow);
                    Row currentRow = table.Rows.Last();

                    for (int j = 0; j < tagsValues.Length; j++)
                    {
                        if (i < tagsValues[j]?.Length)
                            currentRow.ReplaceText(tags[j].Tag, tagsValues[j][i]);
                        else if (i >= tagsValues[j]?.Length)
                            currentRow.ReplaceText(tags[j].Tag, tagsValues[j][tagsValues[j].Length - 1]);
                    }
                }
                table.RemoveRow(patternRowIndex);
            }
            catch
            {

            }
        }

        /// <summary>
        /// Реализация костыля для раздела компетенций.
        /// Заменяет все одинокие теги на параграфы из tagValue, сплитнутого по \n.
        /// </summary>
        /// <param name="doc">Документ DocX</param>
        /// <param name="contentTag">Тег</param>
        /// <param name="lines">Массив строк, которые станут параграфами на месте тега</param>
        protected void ArrayToParagraps(DocX doc, SmartSyllabusTag contentTag, string[] lines)
        {
            try
            {
                // расставим знаки препинания
                if (lines.Length > 0)
                {
                    for (int i = 0; i < lines.Length-1; i++)
                        lines[i] += ";";
                    lines[lines.Length - 1] += ".";
                }

                // добавление параграфов
                string tagString = contentTag.Tag;
                foreach (Paragraph paragraph in doc.Paragraphs)
                {
                    if (paragraph.Text.Trim() == tagString)
                    {
                        Paragraph p = paragraph.InsertParagraphAfterSelf(paragraph);
                        for (int i = 0; i < lines.Length; i++)
                        {
                            p.ReplaceText(tagString, lines[i]);
                            if (i != lines.Length - 1)
                                p = p.InsertParagraphAfterSelf(paragraph);

                        }
                        paragraph.Remove(false);
                    }
                }
            }
            catch
            {

            }
        }
    }
}
