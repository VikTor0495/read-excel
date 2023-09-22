package utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

public class FileUtils {
    
    public static FileInputStream openExcelToSheet(String path, int sheet) throws FileNotFoundException {
        return new FileInputStream(new File("C:\\demo\\student.xls"));
    }
}
