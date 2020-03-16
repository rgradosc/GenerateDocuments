namespace Ayd.AsposeWord.Library
{
    using Ayd.AsposeWord.Library.Enums;
    using System;
    using System.IO;

    public static class DirectoryFileValidator
    {
        public static bool FileExist(this string fullPath)
        {
            return File.Exists(fullPath);
        }

        public static bool DirectoryExist(this string directory)
        {
            return Directory.Exists(directory);
        }

        public static int CreateDirectory(this string directory)
        {
            try
            {
                Directory.CreateDirectory(directory);

                return (int)TypesEvent.SuccessProccess;
            }
            catch (UnauthorizedAccessException)
            {
                return (int)TypesEvent.UnauthorizedAccess;
            }
            catch (ArgumentNullException)
            {
                return (int)TypesEvent.ParamNullOrEmpty;
            }
            catch (Exception)
            {
                return (int)TypesEvent.ErrorGeneric;
            }
        }
    }
}
