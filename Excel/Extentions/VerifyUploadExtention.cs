using Excel.Models.ViewModels;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Excel.Extentions
{
    public static class VerifyUploadExtention
    {
        public static ValidationViewModel IsValidFile(this string path)
        {
            var result = new ValidationViewModel { FileExtention = "", IsValid = false };
            string fileExtention = Path.GetExtension(path);

            switch (fileExtention)
            {
                case (".csv"):
                    result = new ValidationViewModel { FileExtention = ".csv", IsValid = true };
                    break;
                case (".xls"):
                    result = new ValidationViewModel { FileExtention = ".xls", IsValid = true };
                    break;
                case (".xlsx"):
                    result = new ValidationViewModel { FileExtention = ".xlsx", IsValid = true };
                    break;
                case (".chv"):
                    result = new ValidationViewModel { FileExtention = ".chv", IsValid = true };
                    break;
            }

            return result;
        }

        public static Dictionary<string, string> GetDisplayNameList<T>()
        {
            var info = TypeDescriptor.GetProperties(typeof(T))
                .Cast<PropertyDescriptor>()
                .Where(p => p.Attributes.Cast<Attribute>().Any(a => a.GetType() == typeof(RequiredAttribute)))
                .ToDictionary(p => p.Name, p => p.DisplayName);
            return info;
        }
    }
}
