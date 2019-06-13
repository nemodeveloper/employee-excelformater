import java.io.File;

public class EmployeePassProcessor
{
    private final String fileName;
    private final EmployeePassFormatter employeePassFormatter;

    public EmployeePassProcessor(String fileName)
    {
        this.fileName = fileName;
        this.employeePassFormatter = new EmployeePassFormatter(new EmployeePassExcelParser(fileName));
    }

    public void process()
    {
        try
        {
            employeePassFormatter.format();
            if (new File(fileName).delete())
            {
                System.out.println(String.format("Файл %s удален!", fileName));
            }
        }
        catch (Exception e)
        {
            System.out.println("Произошла ошибка форматирования файла!");
            e.printStackTrace();
        }
    }
}
