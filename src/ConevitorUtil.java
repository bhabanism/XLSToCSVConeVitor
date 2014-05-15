import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.util.Iterator;
import java.util.Properties;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ConevitorUtil {
	private static String dataDirName;
	private static String fileType;
	private static String delimiter;

	private static final String DIR_PROP_NAME = "DIR";
	private static final String FILE_TYPE_PROP_NAME = "FILE_TYPE";
	private static final String DELIMITER_PROP_NAME = "DELIMITER";

	static void xlsToText(File inputFile) throws Exception {
		
		File outputFile = new File(getOutputFile(inputFile.getAbsolutePath()));
		System.out.println(outputFile.getAbsolutePath()+"\n\n");
		
		StringBuilder data = new StringBuilder();

		FileOutputStream fos = new FileOutputStream(outputFile);


		Workbook wBook = new HSSFWorkbook(new FileInputStream(inputFile));
		// Get first sheet from the workbook
		Sheet sheet = wBook.getSheetAt(0);
		
		Row row;
		Cell cell;
		// Iterate through each rows from first sheet
		Iterator<Row> rowIterator = sheet.iterator();

		while (rowIterator.hasNext()) {
			row = rowIterator.next();

			// For each row, iterate through each columns
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {

				cell = cellIterator.next();

				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_BOOLEAN:
					data.append(cell.getBooleanCellValue() + delimiter);

					break;
				case Cell.CELL_TYPE_NUMERIC:
					data.append(cell.getNumericCellValue() + delimiter);

					break;
				case Cell.CELL_TYPE_STRING:
					data.append(cell.getStringCellValue() + delimiter);
					break;

				case Cell.CELL_TYPE_BLANK:
					data.append("" + delimiter);
					break;
				default:
					data.append(cell + delimiter);

				}
			}
            data.setLength(data.length()-delimiter.length());
			data.append("\r\n");
		}

		fos.write(data.toString().getBytes());
		fos.close();		
	}
	
	private static String getOutputFile(String inputFileName) throws Exception {
		if(inputFileName.lastIndexOf('.')>0)
        {
           // get last index for '.' char
           int lastIndex = inputFileName.lastIndexOf('.');
           
           // get extension
           String extn = inputFileName.substring(lastIndex);
           String fileName = inputFileName.substring(0,lastIndex);
           // match path name extension
           if(extn.equals(".xls"))
           {
           		return (fileName+"."+fileType);              
           }
        }
		return null;
	}

	private static File[] loadXLSFiles(File dataDir) throws Exception { 

		 FilenameFilter fileNameFilter = new FilenameFilter() {
           @Override
           public boolean accept(File dir, String name) {
              if(name.lastIndexOf('.')>0)
              {
                 // get last index for '.' char
                 int lastIndex = name.lastIndexOf('.');
                 
                 // get extension
                 String str = name.substring(lastIndex);
                 
                 // match path name extension
                 if(str.equals(".xls"))
                 {
                    return true;
                 }
              }
              return false;
           }
        };
        
        return  dataDir.listFiles(fileNameFilter);
	}

	// testing the application
    public static void runConevitor() throws Exception {
		
		loadProperties();
		File dataDir = new File(dataDirName);
		if(dataDir.isDirectory()) {
			
			File[] fList = loadXLSFiles(dataDir);
			System.out.println("dataDir::"+dataDir);
			if(fList.length>0) {
				for(File inputFile : fList) {					
					xlsToText(inputFile);					
				}
				System.out.println("Convertion Complete!");
			} else {
				System.err.println("There are no XLS files to convert!");
			}
		} else {
			System.err.println("Can't find you folder - "+dataDirName);
		}
		
		
	}

	private static void loadProperties()  throws Exception {
		// TODO Auto-generated method stub
		Properties propFile = new Properties();
		propFile.load(new FileInputStream("dir.properties"));
		dataDirName = propFile.getProperty(DIR_PROP_NAME);
		fileType = propFile.getProperty(FILE_TYPE_PROP_NAME);
		delimiter = propFile.getProperty(DELIMITER_PROP_NAME);
	}
}
