
namespace Ayd.AsposeWord.Library
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Linq;
    using Aspose.Words;
    using Aspose.Words.Fields;
    using Enums;
    using Entity;

    public static class DocumentProcessor
    {
        public static int NumberFieldsWithoutValue { get; private set; }

        public static FormFieldCollection GetFormFieldsOfDocument(this Document document)
        {
            FormFieldCollection formFields = document.Range.FormFields;
            return formFields;
        }

        public static void SetValueFieldsInDocument(this Document document, List<FieldProperties> listValues, FormFieldCollection formFields)
        {
            int count = 0;

            foreach (var field in formFields)
            {
                var value = listValues.Where(l => l.Name == field.Name).FirstOrDefault();

                if (value != null)
                {
                    field.Result = value.Result;
                }
                else
                {
                    count++;
                }
            }

            NumberFieldsWithoutValue = count;
        }

        public static int SaveDocument(this Document document, string directoryOutput, string fileNameOutput, string format = ".docx")
        {

            int result = 0;
            try
            {
                var directory = directoryOutput.Substring(directoryOutput.Length - 1, 1) != "\\" ? directoryOutput = directoryOutput + "\\" : directoryOutput;

                string path = $"{directory}{fileNameOutput}{format}";

                switch (format)
                {
                    case ".tiff":
                        document.Save(path, SaveFormat.Tiff);
                        result = (int)TypesEvent.SuccessProccess;
                        break;
                    case ".pdf":
                        document.Save(path, SaveFormat.Pdf);
                        result = (int)TypesEvent.SuccessProccess;
                        break;
                    case ".doc":
                        document.Save(path, SaveFormat.Doc);
                        result = (int)TypesEvent.SuccessProccess;
                        break;
                    case ".docx":
                        document.Save(path, SaveFormat.Docx);
                        result = (int)TypesEvent.SuccessProccess;
                        break;
                    default:
                        result = (int)TypesEvent.FormatNotSupported;
                        break;
                }

                return result;
            }
            catch
            {
                result = (int)TypesEvent.ErrorGeneric;
                return result;
            }
        }

        public static Document RemoveFieldEmpty(this Document document, string fullPathDirectoryOutput, string fileNameOutput, string formatOuput)
        {
            var campos = document.Range.Fields.Where(f => f.Type == FieldType.FieldMergeField).ToList();

            if (campos != null)
            {
                for (int i = 0; i < campos.Count; i++)
                {
                    var campo = campos[i];

                    if (string.IsNullOrEmpty(campo.Result) || (campo.Result.Contains("«") && campo.Result.Contains("»")))
                    {
                        campo.Remove();
                    }
                }

                document.Save($"{fullPathDirectoryOutput}{fileNameOutput}{formatOuput}", SaveFormat.Docx);
            }

            return document;
        }

        public static Field FindFieldByKeyWord(this Document document, string keyWord)
        {
            var fieldSignatures = document.Range.Fields.Where(f => f.Type == FieldType.FieldMergeField && f.DisplayResult.ToLower().Contains(keyWord.ToLower())).FirstOrDefault();

            return fieldSignatures;
        }

        public static FormField FindFormFieldByKeyWord(this Document document, string keyWord)
        {
            var fieldSignatures = document.Range.FormFields.Where(f => f.Type == FieldType.FieldFormTextInput && f.Name.ToLower().Contains(keyWord.ToLower())).FirstOrDefault();

            return fieldSignatures;
        }

        public static void InsertImageInDocument(this Document document, string imageUrl, string fieldName, double width, double height)
        {
            DocumentBuilder documentBuilder = new DocumentBuilder(document);
            documentBuilder.MoveToMergeField(fieldName);
            documentBuilder.InsertImage(Image.FromFile(imageUrl), width, height);
        }

        public static void InsertSignaturFooter(Field field, string value)
        {
            field.Result = value;
        }

        public static string GetFieldNameInMergedField(string displayResult)
        {
            return displayResult.Remove(0, 1).Remove(displayResult.Length - 2, 1);
        }

        public static bool FindField(this Document document, string fieldName)
        {
            if (document != null)
            {
                var field = document.Range.Fields.Where(f => f.Type == FieldType.FieldMergeField && f.DisplayResult.Contains(fieldName)).ToList().Count;

                return Convert.ToBoolean(field);
            }

            return false;
        }

        public static void RotateImage(Bitmap bitmap, string fileName)
        {
            bitmap.RotateFlip(RotateFlipType.Rotate90FlipY);
            bitmap.Save(fileName);
        }
    }
}
