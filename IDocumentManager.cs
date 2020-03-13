namespace Ayd.AsposeWord.Library
{
    public interface IDocumentManager
    {
        /// <summary>
        /// Genera un nuevo documento a partir de un documento base o plantilla, inserta los valores especificados en los campos de texto editables y luego guarda el documento en formato .docx.
        /// </summary>
        /// <param name="pathFileTemplate">Indica la ruta completa donde se encuentra el documento base o plantilla.</param>
        /// <param name="fileNameOutput">Indica el nombre con el cual se guarda el documento final.</param>
        /// <param name="pathDirectoryOutput">Indica la ubicación del directorio donde se guardará el documento final.</param>
        /// <param name="values">Indica la lista de claves y valores necesarios para generar el documento.</param>
        /// <param name="stopProcess">Indica si el proceso se detiene o no al encontrar errores en el proceso.</param>
        /// <returns>Devuelve un entero positivo o negatico que indica si el proceso se completo o no.</returns>
        int GenerateDocument(string pathFileTemplate, string fileNameOutput, string pathDirectoryOutput, string[,] values, bool stopProcess = false);

        /// <summary>
        /// Agregar o inserta un código de barras en el documento especificado.
        /// La plantilla para generar el documento debe contener un campo que en su nombre contenga la palabra barcode.
        /// </summary>
        /// <param name="pathFileTemplate">Indica la ruta completa del documento base o plantilla para insertar el código de barras</param>
        /// <param name="fileNameOutput">Indica el nombre que se le dara al documento final.</param>
        /// <param name="pathDirectoryOutput">Indica la ruta del directorio de salida donde se almacena el documento final.</param>
        /// <param name="pathFileImage">Indica la ruta de ubicación de la imagen.</param>
        /// <param name="width">Indica el ancho de la imagen.</param>
        /// <param name="height">Indica el alto de la imagen.</param>
        /// <param name="formatOutput">Indica el formato de salida del documento final, debe incluirse el punto.</param>
        /// <returns>Devuelve un entero positivo o negativo que indica si el proceso se completo o no.</returns>
        int AddBarCodeInDocument(string pathFileTemplate, string fileNameOutput, string pathDirectoryOutput, string pathFileImage, double width, double height, string formatOutput);

        /// <summary>
        /// Agregar o inserta la imagen de la firma en el documento especificado.
        /// </summary>
        /// <param name="pathFileTemplate">Indica la ruta completa del documento base o plantilla para insertar la imagen.</param>
        /// <param name="fileNameOutput">Indica el nombre que se le dara al documento final.</param>
        /// <param name="pathDirectoryOutput">Indica la ruta del directorio de salida donde se almacena el documento final.</param>
        /// <param name="pathFileImage">Indica la ruta de ubicación de la imagen.</param>
        /// <param name="fieldName">Indica el nombre del campo donde se insertará la firma.</param>
        /// <param name="width">Indica el ancho de la imagen.</param>
        /// <param name="height">Indica el alto de la imagen.</param>
        /// <param name="formatOutput">Indica el formato de salida del documento final, debe incluirse el punto.</param>
        /// <returns>Devuelve un entero positivo o negativo que indica si el proceso se completo o no.</returns>
        int AddSignatureInDocument(string pathFileTemplate, string fileNameOutput, string pathDirectoryOutput, string pathFileImage, string fieldName, double width, double height, string formatOutput);

        /// <summary>
        /// Exporta un documento al formato destino especificado.
        /// </summary>
        /// <param name="pathOriginFile">Indica la ruta completa de la ubicación del documento origen.</param>
        /// <param name="pathDirectoryOutput">Indica la ruta completa de la ubicación del directorio de salida para el documento exportado.</param>
        /// <param name="fileNameOutput">Indica el nombre del archivo de salida sin especificar el formato.</param>
        /// <param name="formatOuput">Indica el formato al que será exportado el documento de origen.</param>
        /// <returns>Retorna un número entero positivo o negativo que indica si se completo o no el proceso.</returns>
        int ExportDocumentToFormat(string pathOriginFile, string pathDirectoryOutput, string fileNameOutput, string formatOuput);
        
        /// <summary>
        /// Genera un documento de solo lectura y lo protege con la contraseña especificada.
        /// </summary>
        /// <param name="pathOriginFile">Indica la ruta completa del documento a proteger.</param>
        /// <param name="pathDirectoryOutput">Indica la ubicación del directorio donde se guardará el documento final.</param>
        /// <param name="fileNameOutput">Indica el nombre con el cual se guarda el documento final.</param>
        /// <param name="password">Indica la palabra clave o contraseña.</param>
        /// <returns>Devuelve un entero positivo o negativo que indica si el proceso se completo o no.</returns>
        int ProtectedDocument(string pathOriginFile, string pathDirectoryOutput, string fileNameOutput, string password);

        /// <summary>
        /// Busca un campo en el documento especificado.
        /// </summary>
        /// <param name="fieldSignature">Indica el nombre del campo a buscar.</param>
        /// <param name="fullPathDocument">Indica la ruta completa del documento en el cual se realiza la busqueda.</param>
        /// <returns>Devuelve verdadero o false si el campo existe o no en el documento.</returns>
        bool FindSignature(string fieldSignature, string fullPathDocument);
    }
}
