import java.util.Calendar;

public class EmployeePass
{
    public enum Direction
    {
        IN,
        OUT;

        public static Direction fromString(String value1, String value2)
        {
            if ("база".equals(value1))
                return OUT;

            if ("база".equals(value2))
                return IN;

            throw new IllegalArgumentException();
        }
    }

    public String fio;
    public Calendar date;
    public Direction direction;
    public String inObject;
    public String outObject;
    public String workType;
}
