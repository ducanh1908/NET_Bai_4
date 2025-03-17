namespace Common
{
    public class Common
    {
        public static string GetValidInput(string fieldName)
        {
            string input;
            do
            {
                Console.Write($"{fieldName}: ");
                input = Console.ReadLine()?.Trim();
                if (string.IsNullOrEmpty(input))
                {
                    Console.WriteLine($"{fieldName} không được để trống!");
                }
            } while (string.IsNullOrEmpty(input));
            return input;
        }

        public static DateTime GetValidDate(string fieldName)
        {
            DateTime date;
            do
            {
                Console.Write($"{fieldName}: ");
                if (DateTime.TryParseExact(Console.ReadLine(), "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out date))
                {
                    return date;
                }
                Console.WriteLine("Ngày không hợp lệ, vui lòng nhập lại!");
            } while (true);
        }

        public static double GetValidDouble(string fieldName)
        {
            double value;
            do
            {
                Console.Write($"{fieldName}: ");
                if (double.TryParse(Console.ReadLine(), out value) && value > 0)
                {
                    return value;
                }
                Console.WriteLine($"{fieldName} không hợp lệ, vui lòng nhập lại!");
            } while (true);
        }
    }
}
