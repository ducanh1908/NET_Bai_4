namespace ConsoleApp
{
    using DataAccess;
    internal class Program
    {
        static void Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            while (true)
            {
                Console.WriteLine("\nMenu:");
                Console.WriteLine("1. Nhập danh sách nhân viên từ bàn phím");
                Console.WriteLine("2. Nhập danh sách nhân viên từ file Excel");
                Console.WriteLine("3. Hiển thị danh sách nhân viên");
                Console.WriteLine("4. Tìm kiếm nhân viên theo thâm niên");
                Console.WriteLine("5. Thoát");
                Console.WriteLine("6. Xuất danh sách nhân viên ra file Excel");
                Console.Write("Chọn một tùy chọn: ");

                switch (Console.ReadLine())
                {
                    case "1":
                        DataAccess.InputEmployees();
                        break;
                    case "2":
                        DataAccess.ImportEmployeesFromExcel();
                        break;
                    case "3":
                        DataAccess.ShowListEmployees();
                        break;
                    case "4":
                        DataAccess.FindEmployeesBySeniority();
                        break;
                    case "6":
                        DataAccess.ExportEmployeesToExcel();
                        break;
                    case "5":
                        return;
                    default:
                        Console.WriteLine("Lựa chọn không hợp lệ!");
                        break;
                }
            }
        }
    }
}
