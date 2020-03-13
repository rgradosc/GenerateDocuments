namespace Ayd.AsposeWord.Library.Mappers
{
    using System.Collections.Generic;
    using Entity;

    public static class FieldDocumentMapper
    {
        public static List<FieldProperties> ConvertArrayToFieldDocument(this string[,] input)
        {
            List<FieldProperties> list = new List<FieldProperties>();

            if (input != null && input.Length > 0)
            {
                var lenght = input.Length / 2;
                string fieldKey = string.Empty, fieldValue = string.Empty;

                for (int i = 0; i < lenght; i++)
                {
                    for (int j = 0; j < 2; j++)
                    {
                        if (j == 0)
                        {
                            fieldKey = input[i, j];
                        }
                        else
                        {
                            fieldValue = input[i, j];
                        }
                    }

                    list.Add(new FieldProperties()
                    {
                        Name = fieldKey,
                        Result = fieldValue,
                    });
                }
            }

            return list;
        }
    }
}
