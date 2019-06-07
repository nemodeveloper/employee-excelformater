import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

public class EmployeePassFormatter
{
    private final SimpleDateFormat dateFormat;
    private final Sheet existSheet;
    private final List<EmployeePass> employeePassList;

    private Map<String, List<EmployeePass>> formatData;

    public EmployeePassFormatter(EmployeePassExcelParser employeePassExcelParser)
    {
        this.dateFormat = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss");
        this.existSheet = employeePassExcelParser.getParsedSheet();
        this.employeePassList = employeePassExcelParser.parse();
    }

    public void format()
    {
        if (formatData == null)
            formatData = formatData();

        try
        {
            writeToFile(buildWorkbook(), buildFileName());
        }
        catch (Exception e)
        {
           throw new RuntimeException(e);
        }
    }

    private Map<String, List<EmployeePass>> formatData()
    {
        Map<String, List<EmployeePass>> result = new TreeMap<>();
        employeePassList.stream()
                .collect(Collectors.groupingBy(employeePass -> employeePass.fio, Collectors.toList()))
                .forEach((s, entities) ->
                {
                    List<EmployeePass> temp = new ArrayList<>();
                    temp.add(entities.stream()
                            .sorted(Comparator.comparing(o -> o.date))
                            .filter(employeePass -> employeePass.direction == EmployeePass.Direction.IN)
                            .findFirst().orElse(null));
                    temp.add(entities.stream()
                            .sorted(Comparator.comparing(o -> o.date))
                            .filter(employeePass -> employeePass.direction == EmployeePass.Direction.OUT)
                            .reduce((first, second) -> second).orElse(null));
                    temp = temp.stream().filter(Objects::nonNull).collect(Collectors.toList());

                    result.put(s, temp.stream().sorted(Comparator.comparing(o -> o.date)).collect(Collectors.toList()));
                });

        return result;
    }

    private String buildFileName()
    {
       return Main.getBasePath() + "access_zone_formatted.xls";
    }

    private Workbook buildWorkbook()
    {
        Workbook book = new HSSFWorkbook();
        Sheet sheet = book.createSheet("Worksheet");

        CellStyle dataCellStyle = book.createCellStyle();
        fillDataCellStyle(dataCellStyle);

        sheet.createRow(0).createCell(0).setCellValue(existSheet.getRow(0).getCell(0).getStringCellValue());
        sheet.getRow(0).getCell(0).setCellStyle(dataCellStyle);
        sheet.createRow(1).createCell(0).setCellValue(existSheet.getRow(1).getCell(0).getStringCellValue());
        sheet.getRow(1).getCell(0).setCellStyle(dataCellStyle);
        sheet.createRow(2).createCell(0).setCellValue(existSheet.getRow(2).getCell(0).getStringCellValue());
        sheet.getRow(2).getCell(0).setCellStyle(dataCellStyle);

        Row headerRow = sheet.createRow(4);
        Row existHeaderRow = existSheet.getRow(3);

        headerRow.createCell(0).setCellValue(existHeaderRow.getCell(0).getStringCellValue());
        headerRow.getCell(0).setCellStyle(dataCellStyle);
        fillHeaderCellStyle(headerRow.getCell(0).getCellStyle(), existHeaderRow.getCell(0).getCellStyle());
        sheet.setColumnWidth(0, existSheet.getColumnWidth(0));

        headerRow.createCell(1).setCellValue(existHeaderRow.getCell(1).getStringCellValue());
        headerRow.getCell(1).setCellStyle(dataCellStyle);
        fillHeaderCellStyle(headerRow.getCell(1).getCellStyle(), existHeaderRow.getCell(1).getCellStyle());
        sheet.setColumnWidth(1, existSheet.getColumnWidth(1));

        headerRow.createCell(2).setCellValue(existHeaderRow.getCell(3).getStringCellValue());
        headerRow.getCell(2).setCellStyle(dataCellStyle);
        fillHeaderCellStyle(headerRow.getCell(2).getCellStyle(), existHeaderRow.getCell(3).getCellStyle());
        sheet.setColumnWidth(2, existSheet.getColumnWidth(3));

        headerRow.createCell(3).setCellValue(existHeaderRow.getCell(2).getStringCellValue());
        headerRow.getCell(3).setCellStyle(dataCellStyle);
        fillHeaderCellStyle(headerRow.getCell(3).getCellStyle(), existHeaderRow.getCell(2).getCellStyle());
        sheet.setColumnWidth(3, existSheet.getColumnWidth(2));

        headerRow.createCell(4).setCellValue(existHeaderRow.getCell(6).getStringCellValue());
        headerRow.getCell(4).setCellStyle(dataCellStyle);
        fillHeaderCellStyle(headerRow.getCell(4).getCellStyle(), existHeaderRow.getCell(6).getCellStyle());
        sheet.setColumnWidth(4, existSheet.getColumnWidth(6));

        int index = 5;
        for (Map.Entry<String, List<EmployeePass>> entry : formatData.entrySet())
        {
            for (EmployeePass employeePass : entry.getValue())
            {
                Row row = sheet.createRow(index);
                row.createCell(0).setCellValue(employeePass.fio);
                row.createCell(1).setCellValue(dateFormat.format(employeePass.date.getTime()));
                row.createCell(2).setCellValue(employeePass.inObject);
                row.createCell(3).setCellValue(employeePass.outObject);
                row.createCell(4).setCellValue(employeePass.workType);

                for (int i = 0; i < 5; ++i)
                {
                    row.getCell(i).setCellStyle(dataCellStyle);
                }

                ++index;
            }
        }

        return book;
    }

    private void writeToFile(Workbook workbook, String fileName) throws Exception
    {
        workbook.write(new FileOutputStream(fileName));
        workbook.close();
    }

    private void fillHeaderCellStyle(CellStyle newCellStyle, CellStyle existCellStyle)
    {

    }

    private void fillDataCellStyle(CellStyle cellStyle)
    {
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);

        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
    }

}
