import com.cemgunduz.model.ExcelReadResponse;
import com.cemgunduz.util.ExcelReadUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * Created by cgunduz on 2/7/15.
 */
public class AppTest {

    public static void main(String args[]) throws IOException, InvalidFormatException {

        // Instantiating the file
        File file = new File("D:\\Developer\\idea-workspace\\excelread\\src\\test\\resource\\test_excel.xlsx");

        // Actually reading, validating and mapping the response, on a excel read resposne wrapper
        ExcelReadResponse excelReadResponse = ExcelReadUtil.getEntityListFromExcelFile(file,Readable.class);

        // Entity list is extracted
        List<Readable> readableList = excelReadResponse.getEntityList();

        // Printing using etc..
        for(Readable readable : readableList)
        {
            System.out.println(readable.getName());
            // ...
        }
    }
}
