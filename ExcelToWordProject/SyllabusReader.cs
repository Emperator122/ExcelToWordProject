using ExcelDataReader;
using ExcelToWordProject.Models;
using ExcelToWordProject.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace ExcelToWordProject
{
    class SyllabusReader : IDisposable
    {     
        public string FilePath { get; }
        public DataSet ExcelData; // ExcelData.Tables["Name"].Rows[11][0] - пример
        SyllabusParameters Parameters;
        IExcelDataReader excelStream;
        FileStream fileStream;



        public SyllabusReader(string filePath, SyllabusParameters parameters)
        {
            FilePath = filePath;
            Parameters = parameters;
            OpenStreams();
           
        }

        public void CloseStreams()
        {
            if (excelStream != null)
                excelStream.Close();
            if (fileStream != null)
                fileStream.Close();
            ExcelData = null;
        }

        public void OpenStreams()
        {
            fileStream = File.Open(FilePath, FileMode.Open, FileAccess.Read);
            excelStream = ExcelReaderFactory.CreateReader(fileStream);
            ExcelData = excelStream.AsDataSet();
        }

        public void Dispose()
        {
            CloseStreams();
        }

        string ExtractDefaultTagData(DefaultSyllabusTag tag)
        {
            try
            {
                string result = ExcelData.Tables[tag.ListName].Rows[tag.RowIndex][tag.ColumnIndex] as string ?? "";
                return result;
            }
            catch
            {
                return "Error<EDTD>";
            }
        }

        public void ConvertToDocx(string resultFolderPath, string baseDocumentPath, string fileNamePrefix = "", IProgress<int> progress = null)
        {
            bool hasActiveSmartTags = Parameters.Tags.FindIndex(el => el is SmartSyllabusTag && el.Active) != -1;
            DocX baseDocument = DocX.Load(baseDocumentPath);

            if (hasActiveSmartTags)
            {
                // Есть ли смарт-теги, работающие со списком компетенций
                bool hasSmartModulesContentTags =
                    Parameters.Tags.FindIndex(el => 
                    el is SmartSyllabusTag && el.ListName == Parameters.ModulesContentListName) != -1;

                // Есть ли смарт-теги, работающие с планом
                bool hasSmartPlanListTags =
                    Parameters.Tags.FindIndex(el =>
                    el is SmartSyllabusTag && el.ListName == Parameters.PlanListName) != -1;

                // получаем все модули
                List<Module> modules = GetAllModules();
                int i = 0;
                foreach(Module module in modules)
                {
                    string safeName = PathUtils.RemoveIllegalFileNameCharacters(fileNamePrefix + module.Name + ".docx");
                    string resultFilePath = Path.Combine(resultFolderPath, safeName);
                    baseDocument = DocX.Load(baseDocumentPath);
                    baseDocument.SaveAs(resultFilePath);
                    DocX doc = DocX.Load(resultFilePath);
                    
                    // получаем компентенции
                    List<Content> contentList = null;
                    if (hasSmartModulesContentTags)
                        contentList = module.ParseContentList(ExcelData, Parameters);

                    // получаем свойства с листа "план"
                    ModuleProperties properties = null;
                    if (hasSmartPlanListTags)
                        properties = module.ParseProperties(ExcelData, Parameters);

                    // бежим по списку тегов
                    foreach (var tag in Parameters.Tags)
                    {
                        // и заполняем каждый активный тег
                        string tagValue = "err";
                        if (tag.Active)
                        {
                            if (tag is SmartSyllabusTag)
                            {
                                SmartSyllabusTag smartTag = tag as SmartSyllabusTag;
                                tagValue = smartTag.ExtractDataFromModule(module, contentList, properties);
                            }
                            else
                            {
                                tagValue = ExtractDefaultTagData(tag as DefaultSyllabusTag);
                            }
                            doc.ReplaceText("<"+tag.Key+">", tagValue);
                        }                     
                    }
                    doc.Save();
                    doc.Dispose();
                    i++;
                    if (progress != null)
                        progress.Report(i * 100 / modules.Count());
                }
            }
            else // просто заменяем теги
            {
                string resultFilePath = Path.Combine(resultFolderPath, fileNamePrefix + ".docx");
                baseDocument.SaveAs(resultFilePath);
                DocX doc = DocX.Load(resultFilePath);
                // бежим по списку тегов
                foreach (BaseSyllabusTag tag in Parameters.Tags)
                {
                    // и заполняем каждый активный тег
                    if (tag is DefaultSyllabusTag && tag.Active)
                        doc.ReplaceText("<" + tag.Key + ">", ExtractDefaultTagData(tag as DefaultSyllabusTag));

                }
                doc.Save();
                doc.Dispose();
                if (progress != null)
                    progress.Report(100);
            }
        }



        public List<Module> GetAllModules()
        {
            SmartSyllabusTag tag =
                Parameters.Tags.Find(
                    tag_ => tag_ is SmartSyllabusTag && (tag_ as SmartSyllabusTag).Type == SmartTagType.ModuleName) as SmartSyllabusTag;
            List<Module> modules = new List<Module>();

            var rows = ExcelData.Tables[tag.ListName].Rows;

            for (int i = 0; i < rows.Count; i++)
            {
                // Имя модуля
                string moduleName = (rows[i][tag.ColumnIndex] as string) ?? "";

                if (moduleName.Trim() == "")
                    continue;

                // Проверка имени на стоп-слова
                bool containsStopWord = false;
                foreach (string stopWord in Parameters.ModuleNameStopWords)
                {
                    containsStopWord = moduleName.ToLower().Contains(stopWord);
                    if (containsStopWord)
                        break;
                }
                if (containsStopWord)
                    continue;

                // Получим индекс модуля
                // т.к. он может быть смещен на несколько ячеек
                // то придется костылить
                string moduleIndex = "";
                for (int j = tag.ColumnIndex - 1; j >= 0; j--)
                {
                    string val = (rows[i][j] as string) ?? "";
                    if (val.Trim() != "")
                    {
                        moduleIndex = val;
                        break;
                    }
                }

                // Строка, содержащая индексы компетенций
                string contentIndexesStr = rows[i][tag.ColumnIndex + 1] as string ?? "";

                modules.Add(new Module(moduleIndex, moduleName, contentIndexesStr));
            }

            return modules;
        }
    }
}
