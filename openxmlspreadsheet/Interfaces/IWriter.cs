using System.IO;

namespace OpenXmlSpreadsheet.Interfaces
{
    public interface IWriter<T>
    {
        string WriteToFile();
        MemoryStream WriteToMemory();
    }
}
