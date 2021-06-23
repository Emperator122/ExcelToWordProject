using ExcelToWordProject.Syllabus;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace ExcelToWordProject.Utils
{
    static class ConfigManager
    {
        /// <summary>
        /// Путь к конфигу
        /// </summary>
        public static string ConfigPath = Application.StartupPath + "\\config.xml";

        /// <summary>
        /// Получение и десериализация данных конфигурации из файла
        /// </summary>
        /// <returns>Класс с настройками</returns>
        public static SyllabusParameters GetConfigData()
        {
            SyllabusParameters settings;
            try
            {
                using (Stream stream = new FileStream(ConfigPath, FileMode.Open))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(SyllabusParameters));
                    settings = (SyllabusParameters)serializer.Deserialize(stream);
                    // десариализуем словарь
                    settings.planListHeaderNames = settings.tempPlanListHeaderNames
                                        .ToDictionary(i => i.Name, i => i.Value);
                    return settings;
                }
            }
            catch
            {
                settings = new SyllabusParameters(true);
                try
                {
                    SaveConfigData(settings);
                }
                catch
                {
                    return settings;
                }
                return settings;
            }
        }

        /// <summary>
        /// Сериализация и сохранение данных конфигурации
        /// </summary>
        /// <param name="data">Файл с записанными настройками</param>
        public static void SaveConfigData(SyllabusParameters data)
        {
            using (Stream writer = new FileStream(ConfigPath, FileMode.Create))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(SyllabusParameters));
                serializer.Serialize(writer, data);
            }
        }
    }
}
