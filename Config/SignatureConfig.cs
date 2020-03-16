namespace Ayd.AsposeWord.Library.Config
{
    public class SignatureConfig
    {
        /// <summary>
        /// El nombre del campo para la imagen de la firma.
        /// </summary>
        public string FieldSignatureName { get; set; }

        /// <summary>
        /// La ruta completa del documento a firmar.
        /// </summary>
        public string FullPathDocument { get; set; }

        /// <summary>
        /// El contenido del pie de la firma
        /// </summary>
        public string FieldFooterValue { get; set; }

        /// <summary>
        /// El nombre del campo para el pie de firma.
        /// </summary>
        public string FieldFooterName { get; set; }

        /// <summary>
        /// La imagen de la firma digital.
        /// </summary>
        public string FullPathImage { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public double Width { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public double Height { get; set; }
    }
}
