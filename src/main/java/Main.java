import java.io.File;

public class Main
{
    public static void main(String[] args)
    {
        try
        {
            String fileName = "access_zone.xls";

            new EmployeePassFormatter(new EmployeePassExcelParser(getBasePath() + fileName)).format();
        }
        catch (Exception e)
        {
            System.out.println("Произошла ошибка форматирования файла!");
            e.printStackTrace();
        }
    }

    public static String getBasePath()
    {
        try
        {
            File jarFile = new File(Main.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath());
            return jarFile.getParent() + File.separator;
        }
        catch (Exception e)
        {
            throw new RuntimeException(e);
        }

    }
}
