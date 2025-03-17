
namespace DataAccess
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using OfficeOpenXml;
    using Common;
    using System.Globalization;
    using System.Text;

    struct Employee
    {
        public string EmployeeID { get; }
        public string Name { get; }
        public DateTime JoinDate { get; }
        public double SalaryFactor { get; }
        public string Position { get; }

        public Employee(string id, string name, DateTime joinDate, double salaryFactor, string position)
        {
            EmployeeID = id;
            Name = name;
            JoinDate = joinDate;
            SalaryFactor = salaryFactor;
            Position = position;
        }

        public int GetSeniority() { return DateTime.Now.Year - JoinDate.Year; }
    }
    public class DataAccess
    {
        static List<Employee> employees = new();
        public static void InputEmployees()
        {
            Console.Write("Nhập số lượng nhân viên: ");
            if (!int.TryParse(Console.ReadLine(), out int count) || count <= 0)
            {
                Console.WriteLine("Số lượng không hợp lệ!");
                return;
            }
            
            for (int i = 0; i < count; i++)
            {
                Console.WriteLine($"\nNhập thông tin cho nhân viên {i + 1}:");

                string employeeID;
                do
                {
                    employeeID = Common.GetValidInput("Mã nhân viên");

                    if (employees.Any(s => s.EmployeeID == employeeID))
                    {
                        Console.WriteLine($"Mã nhân viên {employeeID} đã tồn tại, vui lòng nhập mã khác!");
                    }

                } while (employees.Any(s => s.EmployeeID == employeeID));

                string name = Common.GetValidInput("Tên nhân viên");
                DateTime joinDate = Common.GetValidDate("Ngày vào công ty (dd/MM/yyyy)");
                string position = Common.GetValidInput("Vị trí công việc");
                double salaryFactor = Common.GetValidDouble("Hệ số lương");

                employees.Add(new Employee(employeeID, name, joinDate, salaryFactor, position));
            }
        }

        public static void ImportEmployeesFromExcel()
        {
            //C:\Users\123\Desktop\dsNhanVien.xlsx
            Console.Write("Nhập đường dẫn file Excel: ");
            string filePath = Console.ReadLine();

            if (!File.Exists(filePath))
            {
                Console.WriteLine("File không tồn tại!");
                return;
            }

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;
            StringBuilder listDongLoi =  new StringBuilder();
            for (int row = 2; row <= rowCount; row++)
            {
                try
                {
                    if (string.IsNullOrEmpty(worksheet.Cells[row, 1].Text) || string.IsNullOrEmpty(worksheet.Cells[row, 2].Text))
                    {
                        continue;
                    }
                    string employeeID = worksheet.Cells[row, 1].Text;
                    string name = worksheet.Cells[row, 2].Text;
                    DateTime joinDate = DateTime.MinValue;
                    if (!string.IsNullOrEmpty(worksheet.Cells[row, 3].Text))
                    {
                        if (!DateTime.TryParseExact(worksheet.Cells[row, 3].Text, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out joinDate))
                        {
                            listDongLoi.Append("Định dạng thời gian sai tại dòng " + row);
                            continue;
                        }
                    }
                    string position = "";
                    if (worksheet.Cells[row, 4].Text != "") 
                         position = worksheet.Cells[row, 4].Text;
                    double salaryFactor = 0;
                    if (!string.IsNullOrEmpty(worksheet.Cells[row, 5].Text))
                    {
                        if (!double.TryParse(worksheet.Cells[row, 5].Text, out salaryFactor))
                        {
                            listDongLoi.Append($"Vui lòng nhập hệ số lương hợp lệ tại dòng {row}");
                            continue;
                        }
                    }
                    employees.Add(new Employee(employeeID, name, joinDate, salaryFactor, position));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lỗi tại dòng {row}: {ex.Message}");
                }
            }
            Console.WriteLine($"{listDongLoi}");
        }

        public static void ShowListEmployees()
        {
            Console.WriteLine("\nDanh sách nhân viên:");
            Console.WriteLine("-----------------------------------------------------------");
            Console.WriteLine("| ID   | Tên                | Ngày vào  | Vị trí  | Hệ số |");
            Console.WriteLine("-----------------------------------------------------------");

            foreach (var emp in employees)
            {
                Console.WriteLine($"| {emp.EmployeeID,-5} | {emp.Name,-18} | {emp.JoinDate:dd/MM/yyyy} | {emp.Position,-8} | {emp.SalaryFactor,4} |");
            }
            Console.WriteLine("-----------------------------------------------------------");
        }

        public static void FindEmployeesBySeniority()
        {
            Console.Write("Nhập số năm thâm niên (5 hoặc 10): ");
            if (!int.TryParse(Console.ReadLine(), out int seniority) || (seniority != 5 && seniority != 10))
            {
                Console.WriteLine("Giá trị không hợp lệ! Vui lòng nhập 5 hoặc 10.");
                return;
            }

            var result = employees.FindAll(emp => emp.GetSeniority() >= seniority);

            if (result.Count > 0)
            {
                Console.WriteLine($"\nNhân viên có thâm niên từ {seniority} năm:");
                Console.WriteLine("-------------------------------------------------");
                Console.WriteLine("| ID   | Tên                | Thâm niên (năm) |");
                Console.WriteLine("-------------------------------------------------");

                foreach (var emp in result)
                {
                    Console.WriteLine($"| {emp.EmployeeID,-5} | {emp.Name,-18} | {emp.GetSeniority(),15} |");
                }
                Console.WriteLine("-------------------------------------------------");
            }
            else
            {
                Console.WriteLine("Không có nhân viên nào đạt tiêu chí.");
            }
        }
        public static void ExportEmployeesToExcel()
        {
            if (employees.Count == 0)
            {
                Console.WriteLine("Danh sách nhân viên trống, không thể xuất file!");
                return;
            }

            string filePath = $"C:\\Users\\spm1\\Downloads\\DanhSachNhanVien_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Danh Sách Nhân Viên");
            worksheet.Cells[1, 1].Value = "Mã NV";
            worksheet.Cells[1, 2].Value = "Tên";
            worksheet.Cells[1, 3].Value = "Ngày vào công ty";
            worksheet.Cells[1, 4].Value = "Vị trí";
            worksheet.Cells[1, 5].Value = "Hệ số lương";
            worksheet.Cells[1, 6].Value = "Thâm niên (năm)";
            for (int i = 0; i < employees.Count; i++)
            {
                var emp = employees[i];
                worksheet.Cells[i + 2, 1].Value = emp.EmployeeID;
                worksheet.Cells[i + 2, 2].Value = emp.Name;
                worksheet.Cells[i + 2, 3].Value = emp.JoinDate.ToString("dd/MM/yyyy");
                worksheet.Cells[i + 2, 4].Value = emp.Position;
                worksheet.Cells[i + 2, 5].Value = emp.SalaryFactor;
                worksheet.Cells[i + 2, 6].Value = emp.GetSeniority();
            }
            worksheet.Cells.AutoFitColumns();
            File.WriteAllBytes(filePath, package.GetAsByteArray());
            Console.WriteLine($"Xuất file thành công: {filePath}");
        }
    }
}
