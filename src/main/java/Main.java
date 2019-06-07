import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

public class Main
{
    private static final SimpleDateFormat dateParser = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss");
    private static final SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss");

    public static void main(String[] args) throws Exception
    {
        Sheet sheet = getSheet("test.xls");

        List<Entity> entityList = parseSheet(sheet);
        Map<String, List<Entity>> result = formatData(entityList);

        writeToFile(sheet, result, "FormatData.xls");
    }

    private static void writeToFile(Sheet existSheet, Map<String, List<Entity>> result, String fileName) throws Exception
    {
        Workbook book = new HSSFWorkbook();
        Sheet sheet = book.createSheet("Worksheet");
        sheet.createRow(0).createCell(0).setCellValue(existSheet.getRow(0).getCell(0).getStringCellValue());
        sheet.createRow(1).createCell(0).setCellValue(existSheet.getRow(1).getCell(0).getStringCellValue());
        sheet.createRow(2).createCell(0).setCellValue(existSheet.getRow(2).getCell(0).getStringCellValue());

        Row headerRow = sheet.createRow(3);
        Row existHeaderRow = existSheet.getRow(3);

        headerRow.createCell(0).setCellValue(existHeaderRow.getCell(0).getStringCellValue());
        headerRow.createCell(1).setCellValue(existHeaderRow.getCell(1).getStringCellValue());
        headerRow.createCell(2).setCellValue(existHeaderRow.getCell(3).getStringCellValue());
        headerRow.createCell(3).setCellValue(existHeaderRow.getCell(2).getStringCellValue());
        headerRow.createCell(4).setCellValue(existHeaderRow.getCell(6).getStringCellValue());

        int index = 5;
        for (Map.Entry<String, List<Entity>> entry : result.entrySet())
        {
            for (Entity entity : entry.getValue())
            {
                Row row = sheet.createRow(index);
                row.createCell(0).setCellValue(entity.fio);
                row.createCell(1).setCellValue(dateFormat.format(entity.date.getTime()));
                row.createCell(2).setCellValue(entity.direction == Entity.Direction.IN ? "база" : "----");
                row.createCell(3).setCellValue(entity.direction == Entity.Direction.OUT ? "база" : "----");
                row.createCell(4).setCellValue(entity.workType);

                ++index;
            }
        }

        book.write(new FileOutputStream(fileName));
        book.close();
    }

    private static Map<String, List<Entity>> formatData(List<Entity> entityList)
    {
        Map<String, List<Entity>> result = new TreeMap<>();
        entityList.stream()
            .collect(Collectors.groupingBy(entity -> entity.fio, Collectors.toList()))
            .forEach((s, entities) ->
            {
                List<Entity> temp = new ArrayList<>();
                temp.add(entities.stream()
                        .sorted(Comparator.comparing(o -> o.date))
                        .filter(entity -> entity.direction == Entity.Direction.IN)
                        .findFirst().orElse(null));
                temp.add(entities.stream()
                        .sorted(Comparator.comparing(o -> o.date))
                        .filter(entity -> entity.direction == Entity.Direction.OUT)
                        .reduce((first, second) -> second).orElse(null));
                temp = temp.stream().filter(Objects::nonNull).collect(Collectors.toList());

                result.put(s, temp.stream().sorted(Comparator.comparing(o -> o.date)).collect(Collectors.toList()));
            });

        return result;
    }

    private static Sheet getSheet(String fileName) throws IOException
    {
        try
        {
            InputStream in = Main.class.getResourceAsStream(fileName);
            HSSFWorkbook wb = new HSSFWorkbook(in);

            return wb.getSheetAt(0);

        }
        catch (IOException e)
        {
            e.printStackTrace();
            throw e;
        }
    }

    private static List<Entity> parseSheet(Sheet sheet) throws Exception
    {
        List<Entity> entityList = new ArrayList<Entity>();
        int index = 1;
        for (Row cells : sheet)
        {
            if (index > 4)
            {
                entityList.add(parseRow(cells));
            }
            ++index;
        }

        return entityList;
    }

    private static Entity parseRow(Row row) throws ParseException
    {
        Entity entity = new Entity();
        entity.fio = row.getCell(0).getStringCellValue();
        entity.date = Calendar.getInstance();
        entity.date.setTime(dateParser.parse(row.getCell(1).getStringCellValue()));
        entity.direction = Entity.Direction.fromString(
                row.getCell(2).getStringCellValue(),
                row.getCell(3).getStringCellValue());
        entity.workType = row.getCell(6).getStringCellValue();

        return entity;
    }
}
