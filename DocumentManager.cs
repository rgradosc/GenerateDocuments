namespace Ayd.AsposeWord.Library
{
    using System;
    using System.IO;
    using Aspose.Words;
    using IronBarCode;
    using Enums;
    using Config;
    using Mappers;

    public class DocumentManager : IDocumentManager
    {
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
        public int AddBarCodeInDocument(string pathFileTemplate, string fileNameOutput, string pathDirectoryOutput, string pathFileImage, double width, double height, string formatOutput = ".docx")
        {
            int output = 0;

            try
            {
                if (!pathFileTemplate.FileExist())
                {
                    output = (int)TypesEvent.FileNotFound;
                    return output;
                }

                Document document = new Document(pathFileTemplate);

                var fieldBarCode = DocumentProcessor.FindFieldByKeyWord(document, "barcode");

                if (fieldBarCode == null)
                {
                    output = (int)TypesEvent.FieldNotFound;
                    return output;
                }

                string fieldName = DocumentProcessor.GetFieldNameInMergedField(fieldBarCode.DisplayResult);
                document.InsertImageInDocument(pathFileImage, fieldName, width, height);
                return document.SaveDocument(pathDirectoryOutput, fileNameOutput, formatOutput);
            }
            catch (UnsupportedFileFormatException)
            {
                output = (int)TypesEvent.UnsupportedFileFormat;
            }
            catch (FileCorruptedException)
            {
                output = (int)TypesEvent.FileCorrupted;
            }
            catch (IOException ioex)
            {
                var message = ioex.Message.ToLower();

                if (message.Contains("porque está siendo utilizado en otro proceso"))
                {
                    output = (int)TypesEvent.DocumentInOtherProccess;
                }
                else
                {
                    output = (int)TypesEvent.ErrorGeneric;
                }
            }
            catch (Exception)
            {
                output = (int)TypesEvent.ErrorGeneric;
            }

            return output;
        }

        public int AddBarCodeInDocument(DocumentConfig documentConfig, BarCodeConfig barCodeConfig)
        {
            int output = 0;

            try
            {
                if (!documentConfig.FullPathTemplate.FileExist())
                {
                    output = (int)TypesEvent.FileNotFound;
                    return output;
                }

                Document document = new Document(documentConfig.FullPathTemplate);

                var fieldBarCode = DocumentProcessor.FindFieldByKeyWord(document, barCodeConfig.FieldBarCodeName);

                if (fieldBarCode == null)
                {
                    output = (int)TypesEvent.FieldNotFound;
                    return output;
                }

                string fieldName = DocumentProcessor.GetFieldNameInMergedField(fieldBarCode.DisplayResult);
                document.InsertImageInDocument(barCodeConfig.FullPathImage, fieldName, barCodeConfig.Width, barCodeConfig.Height);
                return document.SaveDocument(documentConfig.FullPathDirectory, documentConfig.FileName, documentConfig.Format);
            }
            catch (UnsupportedFileFormatException)
            {
                output = (int)TypesEvent.UnsupportedFileFormat;
            }
            catch (FileCorruptedException)
            {
                output = (int)TypesEvent.FileCorrupted;
            }
            catch (IOException ioex)
            {
                var message = ioex.Message.ToLower();

                if (message.Contains("porque está siendo utilizado en otro proceso"))
                {
                    output = (int)TypesEvent.DocumentInOtherProccess;
                }
                else
                {
                    output = (int)TypesEvent.ErrorGeneric;
                }
            }
            catch (Exception)
            {
                output = (int)TypesEvent.ErrorGeneric;
            }

            return output;
        }

        /// <summary>
        /// Agregar o inserta la imagen de la firma en el documento especificado.
        /// </summary>
        /// <param name="documentConfig">Es la configuración del documento.</param>
        /// <param name="signatureConfig">Es la configuración de la firma.</param>
        /// <returns>Devuelve un entero positivo o negativo que indica si el proceso se completo o no.</returns>
        public int AddSignatureInDocument(DocumentConfig documentConfig, SignatureConfig signatureConfig)
        {
            int output = 0;

            try
            {
                if (!documentConfig.FullPathTemplate.FileExist())
                {
                    output = (int)TypesEvent.FileNotFound;
                    return output;
                }

                Document document = new Document(documentConfig.FullPathTemplate);

                var fieldSignatures = DocumentProcessor.FindFieldByKeyWord(document, signatureConfig.FieldSignatureName);

                if (fieldSignatures == null)
                {
                    output = (int)TypesEvent.FieldNotFound;
                    return output;
                }

                string fieldSignature = DocumentProcessor.GetFieldNameInMergedField(fieldSignatures.DisplayResult);
                document.InsertImageInDocument(signatureConfig.FullPathImage, fieldSignature, signatureConfig.Width, signatureConfig.Height);

                if (!string.IsNullOrEmpty(signatureConfig.FieldFooterName))
                {
                    var footerField = DocumentProcessor.FindFormFieldByKeyWord(document, signatureConfig.FieldFooterName);

                    if (footerField == null)
                    {
                        output = (int)TypesEvent.FieldNotFound;
                        return output;
                    }

                    footerField.Result = signatureConfig.FieldFooterValue;
                }

                return document.SaveDocument(documentConfig.FullPathDirectory, documentConfig.FileName, ".docx");
            }
            catch (UnsupportedFileFormatException)
            {
                output = (int)TypesEvent.UnsupportedFileFormat;
            }
            catch (FileCorruptedException)
            {
                output = (int)TypesEvent.FileCorrupted;
            }
            catch (IOException ioex)
            {
                var message = ioex.Message.ToLower();

                if (message.Contains("porque está siendo utilizado en otro proceso"))
                {
                    output = (int)TypesEvent.DocumentInOtherProccess;
                }
                else
                {
                    output = (int)TypesEvent.ErrorGeneric;
                }
            }
            catch (Exception ex)
            {
                output = (int)TypesEvent.ErrorGeneric;
            }

            return output;
        }

        /// <summary>
        /// Exporta un documento al formato destino especificado.
        /// </summary>
        /// <param name="pathOriginFile">Indica la ruta completa de la ubicación del documento origen.</param>
        /// <param name="pathDirectoryOutput">Indica la ruta completa de la ubicación del directorio de salida para el documento exportado.</param>
        /// <param name="fileNameOutput">Indica el nombre del archivo de salida sin especificar el formato.</param>
        /// <param name="formatOuput">Indica el formato al que será exportado el documento de origen.</param>
        /// <returns>Retorna un número entero positivo o negativo que indica si se completo o no el proceso.</returns>
        public int ExportDocumentToFormat(string fullPathTemplate, string fullPathDirectoryOutput, string fileNameOutput, string formatOuput)
        {
            int output = 0;

            if (!fullPathTemplate.FileExist())
            {
                output = (int)TypesEvent.FileNotFound;
                return output;
            }

            if (!fullPathDirectoryOutput.DirectoryExist())
            {
                var result = fullPathDirectoryOutput.CreateDirectory();

                if (result != (int)TypesEvent.SuccessProccess)
                {
                    return result;
                }
            }

            try
            {
                output = new Document(fullPathTemplate)
                    .RemoveFieldEmpty(fullPathDirectoryOutput, fileNameOutput, formatOuput)
                    .SaveDocument(fullPathDirectoryOutput, fileNameOutput, formatOuput);
            }
            catch (UnsupportedFileFormatException)
            {
                output = (int)TypesEvent.UnsupportedFileFormat;
            }
            catch (FileCorruptedException)
            {
                output = (int)TypesEvent.FileCorrupted;
            }
            catch (IOException ioex)
            {
                var message = ioex.Message.ToLower();

                if (message.Contains("porque está siendo utilizado en otro proceso"))
                {
                    output = (int)TypesEvent.DocumentInOtherProccess;
                }
                else
                {
                    output = (int)TypesEvent.ErrorGeneric;
                }
            }
            catch
            {
                output = output = (int)TypesEvent.ErrorGeneric;
            }

            return output;
        }

        /// <summary>
        /// Exporta un documento al formato destino especificado.
        /// </summary>
        /// <param name="config">Indica la configuración para generar el documento.</param>
        /// <returns>Retorna un número entero positivo o negativo que indica si se completo o no el proceso.</returns>
        public int ExportDocumentToFormat(DocumentConfig config)
        {
            if (config == null)
            {
                return (int)TypesEvent.ParamNullOrEmpty;
            }

            return ExportDocumentToFormat(config.FullPathTemplate, config.FullPathDirectory, config.FileName, config.Format);
        }

        /// <summary>
        /// Busca un campo en el documento especificado.
        /// </summary>
        /// <param name="fieldSignature">Indica el nombre del campo a buscar.</param>
        /// <param name="fullPathDocument">Indica la ruta completa del documento en el cual se realiza la busqueda.</param>
        /// <returns>Devuelve verdadero o false si el campo existe o no en el documento.</returns>
        public bool FindSignature(string fieldSignature, string fullPathDocument)
        {
            try
            {
                return new Document(fullPathDocument).FindField(fieldSignature);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Genera un codigo de barras a partir de la data especificada y lo guarda como imagen en formato .png
        /// </summary>
        /// <param name="data">Los datos que se van a codificar.</param>
        /// <param name="fullPathDirectory">El directorio de salida de la imagen.</param>
        /// <returns>Devuelve un número entero positivo o negativo.</returns>
        public int GenerateBarCodeAsPNG(string data, string fullPathDirectory, bool rotate)
        {
            try
            {
                var directory = fullPathDirectory.Substring(fullPathDirectory.Length - 1, 1) != "\\" ? fullPathDirectory = fullPathDirectory + "\\" : fullPathDirectory;

                string fileName = $"{directory}barCode.png";
                var barCode = BarcodeWriter.CreateBarcode(data, BarcodeEncoding.Code128);
                var result = barCode.SaveAsPng(fileName);

                if (rotate)
                {
                    DocumentProcessor.RotateImage(result.ToBitmap(), fileName); 
                }

                return (int)TypesEvent.SuccessProccess;
            }
            catch (IronBarCodeEncodingException encex)
            {
                var message = encex.Message;

                if (message.Contains("Bad character in input"))
                {
                    return (int)TypesEvent.BadCharacter;
                }
                else if (message.Contains("Contents length should be between 1 and 80 characters"))
                {
                    return (int)TypesEvent.ContentLength;
                }
                else
                {
                    return (int)TypesEvent.ErrorGeneric;
                }

            }
            catch (Exception)
            {
                return (int)TypesEvent.ErrorGeneric;
            }
        }

        /// <summary>
        /// Genera un nuevo documento a partir de un documento base o plantilla, inserta los valores especificados en los campos de texto editables y luego guarda el documento en formato .docx.
        /// </summary>
        /// <param name="pathFileTemplate">Indica la ruta completa donde se encuentra el documento base o plantilla.</param>
        /// <param name="fileNameOutput">Indica el nombre con el cual se guarda el documento final.</param>
        /// <param name="pathDirectoryOutput">Indica la ubicación del directorio donde se guardará el documento final.</param>
        /// <param name="values">Indica la lista de claves y valores necesarios para generar el documento.</param>
        /// <param name="stopProcess">Indica si el proceso se detiene o no al encontrar errores en el proceso.</param>
        /// <returns>Devuelve un entero positivo o negatico que indica si el proceso se completo o no.</returns>
        public int GenerateDocument(string urlTemplatePath, string fileNameOutput, string urlDirectoryOutput, string[,] values, bool stopProccess = false)
        {
            int output = 0;

            try
            {
                var list = values.ConvertArrayToFieldDocument();

                if (list.Count <= 0)
                {
                    output = (int)TypesEvent.ListDataEmpty;
                    return output;
                }

                if (!urlTemplatePath.FileExist())
                {
                    output = (int)TypesEvent.FileNotFound;
                    return output;
                }

                if (!urlDirectoryOutput.DirectoryExist())
                {
                    var result = urlDirectoryOutput.CreateDirectory();

                    if (result != (int)TypesEvent.SuccessProccess)
                    {
                        return result;
                    }
                }

                Document document = new Document(urlTemplatePath);

                var formFields = document.GetFormFieldsOfDocument();

                if (formFields == null || formFields.Count <= 0)
                {
                    output = (int)TypesEvent.FieldNotFound;
                    return output;
                }

                document.SetValueFieldsInDocument(list, formFields);

                if (DocumentProcessor.NumberFieldsWithoutValue > 0)
                {
                    if (stopProccess)
                    {
                        output = (int)TypesEvent.FieldWithoutValue;
                        return output;
                    }
                }

                output = document.SaveDocument(urlDirectoryOutput, fileNameOutput);
            }
            catch (FileCorruptedException)
            {
                output = (int)TypesEvent.FileCorrupted;
            }
            catch (UnsupportedFileFormatException)
            {
                output = (int)TypesEvent.UnsupportedFileFormat;
            }
            catch (IOException ioex)
            {
                var message = ioex.Message.ToLower();

                if (message.Contains("porque está siendo utilizado en otro proceso"))
                {
                    output = (int)TypesEvent.DocumentInOtherProccess;
                }
                else
                {
                    output = (int)TypesEvent.ErrorGeneric;
                }
            }
            catch (Exception)
            {
                output = (int)TypesEvent.ErrorGeneric;
            }

            return output;
        }

        /// <summary>
        /// Genera un documento a partir de una clase de configuración.
        /// </summary>
        /// <param name="config">Indica la configuración para generar el documento.</param>
        /// <param name="values">Indica la lista de claves y valores necesarios para generar el documento.</param>
        /// <returns>Devuelve un entero positivo o negatico que indica si el proceso se completo o no.</returns>
        public int GenerateDocument(DocumentConfig config, string[,] values)
        {
            if (config == null)
            {
                return (int)TypesEvent.ParamNullOrEmpty;
            }

            return GenerateDocument(config.FullPathTemplate, config.FileName, config.FullPathDirectory, values, config.StopProcess);
        }

        /// <summary>
        /// Genera un documento en formato .docx de solo lectura y lo protege con la contraseña especificada.
        /// </summary>
        /// <param name="pathOriginFile">Indica la ruta completa del documento a proteger.</param>
        /// <param name="pathDirectoryOutput">Indica la ubicación del directorio donde se guardará el documento final.</param>
        /// <param name="fileNameOutput">Indica el nombre con el cual se guarda el documento final.</param>
        /// <param name="password">Indica la palabra clave o contraseña.</param>
        /// <returns>Devuelve un entero positivo o negativo que indica si el proceso se completo o no.</returns>
        public int ProtectedDocument(string pathOriginFile, string pathDirectoryOutput, string fileNameOutput, string password)
        {
            int output = 0;

            if (!pathOriginFile.FileExist())
            {
                return (int)TypesEvent.FileNotFound;
            }

            if (!pathDirectoryOutput.DirectoryExist())
            {

                var result = pathDirectoryOutput.CreateDirectory();

                if (result != (int)TypesEvent.SuccessProccess)
                {
                    return result;
                }
            }

            if (string.IsNullOrEmpty(password))
            {
                return (int)TypesEvent.PasswordEmpty;
            }

            try
            {
                Document document = new Document(pathOriginFile);
                document.Protect(ProtectionType.ReadOnly, password);
                return document.SaveDocument(pathDirectoryOutput, fileNameOutput, format: ".docx");
            }
            catch (FileCorruptedException)
            {
                output = (int)TypesEvent.FileCorrupted;
            }
            catch (UnsupportedFileFormatException)
            {
                output = (int)TypesEvent.UnsupportedFileFormat;
            }
            catch (IOException ioex)
            {
                var message = ioex.Message.ToLower();

                if (message.Contains("porque está siendo utilizado en otro proceso"))
                {
                    output = (int)TypesEvent.DocumentInOtherProccess;
                }
                else
                {
                    output = (int)TypesEvent.ErrorGeneric;
                }
            }
            catch (Exception)
            {
                output = (int)TypesEvent.ErrorGeneric;
            }

            return output;
        }
    }
}
