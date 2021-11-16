import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.codehaus.jackson.map.ObjectMapper;
import org.json.simple.JSONObject;
public class ExtoJson extends Student
{
    public static void main(String[] args) throws IOException
    {
        List customers = readExcelFile("D:\\new1.xlsx");
        writeObjects2Jsonfile(customers,"D:\\File1.json");
        System.out.println("Done");
    }

    private static List readExcelFile(String filePath)
    {
        try
        {
            FileInputStream excelFile = new FileInputStream(new File(filePath));

            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("Sheet1");
            Iterator rows = sheet.iterator();
            List lstStudents = new ArrayList();
            int rowNumber = 0;
            while (rows.hasNext())
            {
                Row currentRow = (Row) rows.next();
                if (rowNumber == 0)
                {
                    rowNumber++;
                    continue;
                }
                Iterator cellsInRow = currentRow.iterator();
                Student sust = new Student();
                int cellIndex = 0;
                while (cellsInRow.hasNext())
                {
                    Cell currentCell = (Cell) cellsInRow.next();
                    if (cellIndex == 0)
                    {
                        sust.setName(currentCell.getStringCellValue());
                    }
                    else if (cellIndex == 1)
                    {
                        sust.setAge((int) currentCell.getNumericCellValue());
                    }
                    else if (cellIndex == 2)
                    {
                        sust.setTotalMarks((int) currentCell.getNumericCellValue());
                    }
                    cellIndex++;
                }
                lstStudents.add(sust);
            }
            workbook.close();
            return lstStudents;
        }
        catch(IOException e)
            {
                throw new RuntimeException("FAIL!-->message =" + e.getMessage());
            }
        }
        private static void writeObjects2Jsonfile (List Student, String pathFile) throws IOException
        {
            ObjectMapper mapper = new ObjectMapper();
            File file = new File(pathFile);
            try
            {
                mapper.writerWithDefaultPrettyPrinter().writeValue(file, Student);
            }
            catch (IOException e)
            {
                e.printStackTrace();
            }
        }
    }
