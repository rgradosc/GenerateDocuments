namespace Ayd.AsposeWord.Library.Config
{
    public class BarCodeConfig
    {
        /// <summary>
        /// El nombre del campo para la imagen de la firma.
        /// </summary>
        public string FieldBarCodeName { get; set; }

        /// <summary>
        /// La ruta completa del documento a firmar.
        /// </summary>
        public string FullPathDocument { get; set; }

        /// <summary>
        /// La imagen de la firma digital.
        /// </summary>
        public string FullPathImage { get; set; }

        /// <summary>
        /// La anchura de la imagen.
        /// </summary>
        public double Width { get; set; }

        /// <summary>
        /// La altura de la imagen.
        /// </summary>
        public double Height { get; set; }
    }
}
