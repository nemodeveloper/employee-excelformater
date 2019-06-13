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
        if (formatFile())
        {
            removeBaseFile();
        }
    }

    private boolean formatFile()
    {
        try
        {
            employeePassFormatter.format();
            return true;
        }
        catch (Exception e)
        {
            System.out.println("Произошла ошибка форматирования файла!");
            e.printStackTrace();
        }

        return false;
    }

    private void removeBaseFile()
    {
        try
        {
            if (new File(fileName).delete())
            {
                System.out.println(String.format("Файл %s удален!", fileName));
            }
        }
        catch (Exception e)
        {
            System.out.println(String.format("Произошла ошибка удаления файла %s!", fileName));
            e.printStackTrace();
        }
    }
}
