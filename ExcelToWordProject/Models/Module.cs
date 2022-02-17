using ExcelToWordProject.Models;

namespace ExcelToWordProject
{
    public class Module
    {
        public string Index;
        public string Name;
        public string[] ContentIndexes;

        public ModuleProperties Properties;

        public Module(string index, string name, string сontentIndexesStr, char contentDelimiter = ';', ModuleProperties moduleProperties = null)
        {
            Properties = moduleProperties;
            Index = index;
            Name = name;
            ContentIndexes = сontentIndexesStr.Split(contentDelimiter);
            for (int i = 0; i < ContentIndexes.Length; i++)
            {
                ContentIndexes[i] = ContentIndexes[i].Trim();
            }
        }

        public override string ToString()
        {
            return Index + " " + Name;
        }
    }
}
