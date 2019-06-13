import java.io.File;

public class Main
{
    public static void main(String[] args)
    {
        new EmployeePassProcessor(args[0]).process();
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
