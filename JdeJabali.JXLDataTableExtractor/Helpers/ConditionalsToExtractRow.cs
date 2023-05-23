namespace JdeJabali.JXLDataTableExtractor.Helpers
{
    public static partial class ConditionalsToExtractRow
    {
        /// <summary>
        /// The <paramref name="value"/> parameter is always ignored.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool AlwaysTrue(string value)
        {
            return true;
        }

        public static bool IsNotNullOrEmpty(string value)
        {
            return !string.IsNullOrEmpty(value);
        }

        public static bool HasNumericValueAboveZero(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            _ = decimal.TryParse(value, out decimal result);

            return result > 0;
        }
    }
}
