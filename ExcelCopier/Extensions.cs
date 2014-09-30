namespace ExcelCopier
{
    public static class Extensions
    {
        public static string FormatWith(this string str, params object[] args)
        {
            return string.Format(str, args);
        }
    }
}
