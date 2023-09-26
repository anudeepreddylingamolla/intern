using OfficeOpenXml;
using System;

class ExcelFunctionsDemo
{
    static void Main()

    {
        string filePath = "BOOK1.xlsx"; //
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(filePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; 
            // VLOOKUP
            Console.WriteLine("Enter the VLOOKUP value:");
            string lookupValue = Console.ReadLine();

            Console.WriteLine("Enter the lookup column number:");
            int lookupColumn = int.Parse(Console.ReadLine());

            Console.WriteLine("Enter the result column number:");
            int resultColumn = int.Parse(Console.ReadLine());

            object vlookupResult = Vlookup(worksheet, lookupValue, lookupColumn, resultColumn);
            Console.WriteLine($"VLOOKUP Result: {vlookupResult}");

            // HLOOKUP
            Console.WriteLine("Enter the HLOOKUP value:");
            string hlookupValue = Console.ReadLine();

            Console.WriteLine("Enter the lookup row number:");
            int hlookupRow = int.Parse(Console.ReadLine());

            Console.WriteLine("Enter the lookup column number:");
            int hlookupColumn = int.Parse(Console.ReadLine());

            object hlookupResult = Hlookup(worksheet, hlookupValue, hlookupRow, hlookupColumn);
            Console.WriteLine($"HLOOKUP Result: {hlookupResult}");

            // COUNT
            Console.WriteLine("Enter the COUNT column number:");
            int countColumn = int.Parse(Console.ReadLine());

            int count = Count(worksheet, countColumn);
            Console.WriteLine($"COUNT: {count}");

            // CONCATENATE
            Console.WriteLine("Enter the CONCATENATE row number:");
            int concatenateRow = int.Parse(Console.ReadLine());

            Console.WriteLine("Enter the first column number for CONCATENATE:");
            int concatenateColumn1 = int.Parse(Console.ReadLine());

            Console.WriteLine("Enter the second column number for CONCATENATE:");
            int concatenateColumn2 = int.Parse(Console.ReadLine());

            string concatenatedValue = Concatenate(worksheet, concatenateRow, concatenateColumn1, concatenateColumn2);
            Console.WriteLine($"CONCATENATE: {concatenatedValue}");

            // INDEX
            Console.WriteLine("Enter the INDEX row number:");
            int indexRow = int.Parse(Console.ReadLine());

            Console.WriteLine("Enter the INDEX column number:");
            int indexColumn = int.Parse(Console.ReadLine());

            object indexValue = Index(worksheet, indexRow, indexColumn);
            Console.WriteLine($"INDEX Result: {indexValue}");

            // SUM
            Console.WriteLine("Enter the SUM column number:");
            int sumColumn = int.Parse(Console.ReadLine());

            double sum = Sum(worksheet, sumColumn);
            Console.WriteLine($"SUM: {sum}");

            // AVERAGE
            Console.WriteLine("Enter the AVERAGE column number:");
            int averageColumn = int.Parse(Console.ReadLine());

            double average = Average(worksheet, averageColumn);
            Console.WriteLine($"AVERAGE: {average}");

            // MIN
            Console.WriteLine("Enter the MINIMUM column number:");
            int minColumn = int.Parse(Console.ReadLine());

            double minimum = Minimum(worksheet, minColumn);
            Console.WriteLine($"MINIMUM: {minimum}");

            // MAX
            Console.WriteLine("Enter the MAXIMUM column number:");
            int maxColumn = int.Parse(Console.ReadLine());

            double maximum = Maximum(worksheet, maxColumn);
            Console.WriteLine($"MAXIMUM: {maximum}");

            // RANGE
            Console.WriteLine("Enter the RANGE column number:");
            int rangeColumn = int.Parse(Console.ReadLine());

            double range = Range(worksheet, rangeColumn);
            Console.WriteLine($"RANGE: {range}");
        }
    }

    static object Vlookup(ExcelWorksheet worksheet, string lookupValue, int lookupColumn, int resultColumn)
    {
        int startRow = worksheet.Dimension.Start.Row;
        int endRow = worksheet.Dimension.End.Row;

        for (int row = startRow; row <= endRow; row++)
        {
            string cellValue = worksheet.Cells[row, lookupColumn].GetValue<string>();
            if (cellValue == lookupValue)
            {
                return worksheet.Cells[row, resultColumn].Value;
            }
        }

        return null;
    }

    static object Hlookup(ExcelWorksheet worksheet, string lookupValue, int lookupRow, int lookupColumn)
    {
        int startColumn = worksheet.Dimension.Start.Column;
        int endColumn = worksheet.Dimension.End.Column;

        for (int col = startColumn; col <= endColumn; col++)
        {
            string cellValue = worksheet.Cells[lookupRow, col].GetValue<string>();
            if (cellValue == lookupValue)
            {
                return worksheet.Cells[lookupRow

, lookupColumn].Value;
            }
        }

        return null;
    }

    static int Count(ExcelWorksheet worksheet, int countColumn)
    {
        int startRow = worksheet.Dimension.Start.Row;
        int endRow = worksheet.Dimension.End.Row;
        int count = 0;

        for (int row = startRow; row <= endRow; row++)
        {
            if (worksheet.Cells[row, countColumn].Value != null)
            {
                count++;
            }
        }

        return count;
    }

    static string Concatenate(ExcelWorksheet worksheet, int row, int column1, int column2)
    {
        string value1 = worksheet.Cells[row, column1].GetValue<string>();
        string value2 = worksheet.Cells[row, column2].GetValue<string>();

        return string.Concat(value1, value2);
    }

    static object Index(ExcelWorksheet worksheet, int row, int column)
    {
        return worksheet.Cells[row, column].Value;
    }

    static double Sum(ExcelWorksheet worksheet, int sumColumn)
    {
        int startRow = worksheet.Dimension.Start.Row;
        int endRow = worksheet.Dimension.End.Row;
        double sum = 0;

        for (int row = startRow; row <= endRow; row++)
        {
            if (worksheet.Cells[row, sumColumn].Value != null)
            {
                double cellValue = worksheet.Cells[row, sumColumn].GetValue<double>();
                sum += cellValue;
            }
        }

        return sum;
    }

    static double Average(ExcelWorksheet worksheet, int averageColumn)
    {
        int startRow = worksheet.Dimension.Start.Row;
        int endRow = worksheet.Dimension.End.Row;
        double sum = 0;
        int count = 0;

        for (int row = startRow; row <= endRow; row++)
        {
            if (worksheet.Cells[row, averageColumn].Value != null)
            {
                double cellValue = worksheet.Cells[row, averageColumn].GetValue<double>();
                sum += cellValue;
                count++;
            }
        }

        return sum / count;
    }

    static double Minimum(ExcelWorksheet worksheet, int minColumn)
    {
        int startRow = worksheet.Dimension.Start.Row;
        int endRow = worksheet.Dimension.End.Row;
        double minimum = double.MaxValue;

        for (int row = startRow; row <= endRow; row++)
        {
            if (worksheet.Cells[row, minColumn].Value != null)
            {
                double cellValue = worksheet.Cells[row, minColumn].GetValue<double>();
                if (cellValue < minimum)
                {
                    minimum = cellValue;
                }
            }
        }

        return minimum;
    }

    static double Maximum(ExcelWorksheet worksheet, int maxColumn)
    {
        int startRow = worksheet.Dimension.Start.Row;
        int endRow = worksheet.Dimension.End.Row;
        double maximum = double.MinValue;

        for (int row = startRow; row <= endRow; row++)
        {
            if (worksheet.Cells[row, maxColumn].Value != null)
            {
                double cellValue = worksheet.Cells[row, maxColumn].GetValue<double>();
                if (cellValue > maximum)
                {
                    maximum = cellValue;
                }
            }
        }

        return maximum;
    }

    static double Range(ExcelWorksheet worksheet, int rangeColumn)
    {
        double minimum = Minimum(worksheet, rangeColumn);
        double maximum = Maximum(worksheet, rangeColumn);

        return maximum - minimum;
    }
}