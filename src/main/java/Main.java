import java.io.File;
import java.util.Arrays;
import java.util.List;

public class Main
{
    public static void main(String[] args)
    {
        new EmployeePassProcessor(getFileName(args)).process();
    }

    private static String getFileName(String[] args)
    {
        List<String> argList = Arrays.asList(args);
        return argList.size() != 0
                ? argList.get(0)
                : getBasePath() + "access_zone.xls";
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
