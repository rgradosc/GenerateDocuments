namespace Ayd.AsposeWord.Library.Config
{
    public class DocumentConfig
    {
        /// <summary>
        /// Ruta completa del documento base o plantilla.
        /// </summary>
        public string FullPathTemplate { get; set; }

        /// <summary>
        /// Ruta completa del directorio donde se coloca el documento final.
        /// </summary>
        public string FullPathDirectory { get; set; }

        /// <summary>
        /// Nombre para el documento final.
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// Formato al que será generado el documento.
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// Indica si el documento generado es de solo lectura.
        /// </summary>
        public bool IsReadOnly { get; set; }

        /// <summary>
        /// Palabra clave para proteger los documentos de solo lectura.
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// Indica si el proceso se detendrá o no al producirse errores.
        /// </summary>
        public bool StopProcess { get; set; }
    }
}
