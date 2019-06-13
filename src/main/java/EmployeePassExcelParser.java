import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

public class EmployeePassExcelParser
{
    private final String fileName;
    private final SimpleDateFormat dateParser;
    private final Sheet sheet;
    private List<EmployeePass> employeePassList;

    public EmployeePassExcelParser(String fileName)
    {
        this.fileName = fileName;
        this.dateParser = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss");
        this.sheet = getSheet(fileName);
    }

    public List<EmployeePass> parse()
    {
        if (employeePassList == null)
        {
            employeePassList = parseSheet();
        }

        return employeePassList;
    }

    public Sheet getParsedSheet()
    {
        return sheet;
    }

    public String getFileName()
    {
        return fileName;
    }

    private Sheet getSheet(String fileName)
    {
        try
        {
            Workbook workbook = WorkbookFactory.create(new File(fileName));
            Sheet sheet = workbook.getSheetAt(0);
            workbook.close();

            return sheet;
        }
        catch (Exception e)
        {
            throw new RuntimeException(e);
        }
    }

    private List<EmployeePass> parseSheet()
    {
        List<EmployeePass> employeePassList = new ArrayList<EmployeePass>();
        int index = 1;
        for (Row cells : sheet)
        {
            // пропускаем заголовки
            if (index > 4)
            {
                employeePassList.add(parseRow(cells));
            }
            ++index;
        }

        return employeePassList;
    }

    private EmployeePass parseRow(Row row)
    {
        try
        {
            EmployeePass employeePass = new EmployeePass();
            employeePass.fio = row.getCell(0).getStringCellValue();
            employeePass.date = Calendar.getInstance();
            employeePass.date.setTime(dateParser.parse(row.getCell(1).getStringCellValue()));
            employeePass.direction = EmployeePass.Direction.fromString(
                    row.getCell(2).getStringCellValue(),
                    row.getCell(3).getStringCellValue());
            employeePass.inObject = row.getCell(3).getStringCellValue();
            employeePass.outObject = row.getCell(2).getStringCellValue();
            employeePass.workType = row.getCell(6).getStringCellValue();

            return employeePass;
        }
        catch (Exception e)
        {
            throw new RuntimeException("Ошибка парсинга строки файла - " + row.getRowNum(), e);
        }
    }

}
