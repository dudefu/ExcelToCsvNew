package com.ow.util;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.ArrayList;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelToCSV {

    private static Logger logger = LoggerFactory.getLogger(ExcelToCSV.class);

    private Workbook workbook;
    private ArrayList<ArrayList<String>> csvData;
    private int maxRowWidth;
    private int formattingConvention;
    private DataFormatter formatter;
    private FormulaEvaluator evaluator;
    private String separator;

    private static final String CSV_FILE_EXTENSION = ".csv";
    private static final String DEFAULT_SEPARATOR = ",";

    public static final int EXCEL_STYLE_ESCAPING = 0;
    public static final int UNIX_STYLE_ESCAPING = 1;

    public void convertExcelToCSV(String strSource, String strDestination)
            throws IOException,
            IllegalArgumentException {

        this.convertExcelToCSV(strSource, strDestination,
                ExcelToCSV.DEFAULT_SEPARATOR, ExcelToCSV.EXCEL_STYLE_ESCAPING);
    }

    public void convertExcelToCSV(String strSource, String strDestination,
                                  String separator)
            throws IOException,
            IllegalArgumentException {

        this.convertExcelToCSV(strSource, strDestination,
                separator, ExcelToCSV.EXCEL_STYLE_ESCAPING);
    }

    public void convertExcelToCSV(String strSource, String strDestination,
                                  String separator, int formattingConvention)
            throws IOException,
            IllegalArgumentException {
        File source = new File(strSource);
        File destination = new File(strDestination);
        File[] filesList;
        String destinationFilename;
        Sheet sheet;

        //判断文件或者文件夹是否存在
        if(!source.exists()) {
            throw new IllegalArgumentException("The source for the Excel " +
                    "file(s) cannot be found.");
        }

        if(!destination.exists()) {
            throw new IllegalArgumentException("The folder/directory for the " +
                    "converted CSV file(s) does not exist.");
        }
        if(!destination.isDirectory()) {
            throw new IllegalArgumentException("The destination for the CSV " +
                    "file(s) is not a directory/folder.");
        }

        if(formattingConvention != ExcelToCSV.EXCEL_STYLE_ESCAPING &&
                formattingConvention != ExcelToCSV.UNIX_STYLE_ESCAPING) {
            throw new IllegalArgumentException("The value passed to the " +
                    "formattingConvention parameter is out of range.");
        }

        this.separator = separator;
        this.formattingConvention = formattingConvention;

        if(source.isDirectory()) {
            filesList = source.listFiles(new ExcelFilenameFilter());
        }else {
            filesList = new File[]{source};
        }

        if (filesList != null) {
            for(File excelFile : filesList) {

                this.openWorkbook(excelFile);
//                this.convertToCSV();

                String fileName = excelFile.getName();
                String excelFilename = fileName.substring(
                        0, fileName.lastIndexOf("."));

                int numSheets = this.workbook.getNumberOfSheets();
                for (int i = 0; i < numSheets; i++) {
                    sheet = this.workbook.getSheetAt(i);
                    this.convertToCSV(sheet);
                    String sheetName = sheet.getSheetName();
                    destinationFilename = sheet.getSheetName();
                    String destinationFile = destinationFilename + ExcelToCSV.CSV_FILE_EXTENSION;
                    destination = new File(strDestination+"/"+excelFilename+"/"+destinationFilename);
                    if(!destination.isDirectory()){
                        destination.mkdirs();
                    }
                    this.saveCSVFile(new File(destination, destinationFile));
                }
            }
        }
    }

    /**
     * Open an Excel workbook ready for conversion.
     *
     * @param file An instance of the File class that encapsulates a handle
     *        to a valid Excel workbook. Note that the workbook can be in
     *        either binary (.xls) or SpreadsheetML (.xlsx) format.
     * @throws java.io.FileNotFoundException Thrown if the file cannot be located.
     * @throws java.io.IOException Thrown if a problem occurs in the file system.
     */
    private void openWorkbook(File file) throws FileNotFoundException,
            IOException {
        System.out.println("Opening workbook [" + file.getName() + "]");
        try (FileInputStream fis = new FileInputStream(file)) {

            // Open the workbook and then create the FormulaEvaluator and
            // DataFormatter instances that will be needed to, respectively,
            // force evaluation of forumlae found in cells and create a
            // formatted String encapsulating the cells contents.
            this.workbook = WorkbookFactory.create(fis);
            this.evaluator = this.workbook.getCreationHelper().createFormulaEvaluator();
            this.formatter = new DataFormatter(true);
        }
    }

    /**
     * Called to convert the contents of the currently opened workbook into
     * a CSV file.
     */
    private void convertToCSV() {
        Sheet sheet;
        Row row;
        int lastRowNum;
        this.csvData = new ArrayList<>();

        System.out.println("Converting files contents to CSV format.");

        // Discover how many sheets there are in the workbook....
        int numSheets = this.workbook.getNumberOfSheets();

        // and then iterate through them.
        for(int i = 0; i < numSheets; i++) {
            sheet = this.workbook.getSheetAt(i);
            if(sheet.getPhysicalNumberOfRows() > 0) {
                lastRowNum = sheet.getLastRowNum();
                for(int j = 0; j <= lastRowNum; j++) {
                    row = sheet.getRow(j);
                    this.rowToCSV(row);
                }
            }
        }
    }

    /**
     * 自定义方法
     * 功能：转换单个sheet为csv文件
     * @param sheet
     */
    private void convertToCSV(Sheet sheet){
        Row row;
        int lastRowNum;
        this.csvData = new ArrayList<>();

        System.out.println("Converting files contents to CSV format.");

        if(sheet.getPhysicalNumberOfRows() > 0) {
            lastRowNum = sheet.getLastRowNum();
            for(int j = 0; j <= lastRowNum; j++) {
                row = sheet.getRow(j);
                this.rowToCSV(row);
            }
        }
    }

    /**
     * Called to actually save the data recovered from the Excel workbook
     * as a CSV file.
     *
     * @param file An instance of the File class that encapsulates a handle
     *             referring to the CSV file.
     * @throws java.io.FileNotFoundException Thrown if the file cannot be found.
     * @throws java.io.IOException Thrown to indicate and error occurred in the
     *                             underylying file system.
     */
    private void saveCSVFile(File file)
            throws FileNotFoundException, IOException {
        ArrayList<String> line;
        StringBuffer buffer;
        String csvLineElement;

        // Open a writer onto the CSV file.
        try (BufferedWriter bw = new BufferedWriter(new FileWriter(file))) {

            System.out.println("Saving the CSV file [" + file.getName() + "]");

            // Step through the elements of the ArrayList that was used to hold
            // all of the data recovered from the Excel workbooks' sheets, rows
            // and cells.
            for(int i = 0; i < this.csvData.size(); i++) {
                buffer = new StringBuffer();

                // Get an element from the ArrayList that contains the data for
                // the workbook. This element will itself be an ArrayList
                // containing Strings and each String will hold the data recovered
                // from a single cell. The for() loop is used to recover elements
                // from this 'row' ArrayList one at a time and to write the Strings
                // away to a StringBuffer thus assembling a single line for inclusion
                // in the CSV file. If a row was empty or if it was short, then
                // the ArrayList that contains it's data will also be shorter than
                // some of the others. Therefore, it is necessary to check within
                // the for loop to ensure that the ArrayList contains data to be
                // processed. If it does, then an element will be recovered and
                // appended to the StringBuffer.
                line = this.csvData.get(i);
                for(int j = 0; j < this.maxRowWidth; j++) {
                    if(line.size() > j) {
                        csvLineElement = line.get(j);
                        if(csvLineElement != null) {
                            buffer.append(this.escapeEmbeddedCharacters(
                                    csvLineElement));
                        }
                    }
                    if(j < (this.maxRowWidth - 1)) {
                        buffer.append(this.separator);
                    }
                }

                // Once the line is built, write it away to the CSV file.
                bw.write(buffer.toString().trim());

                // Condition the inclusion of new line characters so as to
                // avoid an additional, superfluous, new line at the end of
                // the file.
                if(i < (this.csvData.size() - 1)) {
                    bw.newLine();
                }
            }
        }
    }

    /**
     * Called to convert a row of cells into a line of data that can later be
     * output to the CSV file.
     *
     * @param row An instance of either the HSSFRow or XSSFRow classes that
     *            encapsulates information about a row of cells recovered from
     *            an Excel workbook.
     */
    private void rowToCSV(Row row) {
        Cell cell;
        int lastCellNum;
        ArrayList<String> csvLine = new ArrayList<>();

        // Check to ensure that a row was recovered from the sheet as it is
        // possible that one or more rows between other populated rows could be
        // missing - blank. If the row does contain cells then...
        if(row != null) {

            // Get the index for the right most cell on the row and then
            // step along the row from left to right recovering the contents
            // of each cell, converting that into a formatted String and
            // then storing the String into the csvLine ArrayList.
            lastCellNum = row.getLastCellNum();
            for(int i = 0; i <= lastCellNum; i++) {
                cell = row.getCell(i);
                if(cell == null) {
                    csvLine.add("");
                }
                else {
                    if(cell.getCellType() != CellType.FORMULA) {
                        csvLine.add(this.formatter.formatCellValue(cell));
                    }
                    else {
                        csvLine.add(this.formatter.formatCellValue(cell, this.evaluator));
                    }
                }
            }
            // Make a note of the index number of the right most cell. This value
            // will later be used to ensure that the matrix of data in the CSV file
            // is square.
            if(lastCellNum > this.maxRowWidth) {
                this.maxRowWidth = lastCellNum;
            }
        }
        this.csvData.add(csvLine);
    }

    /**
     * Checks to see whether the field - which consists of the formatted
     * contents of an Excel worksheet cell encapsulated within a String - contains
     * any embedded characters that must be escaped. The method is able to
     * comply with either Excel's or UNIX formatting conventions in the
     * following manner;
     *
     * With regard to UNIX conventions, if the field contains any embedded
     * field separator or EOL characters they will each be escaped by prefixing
     * a leading backspace character. These are the only changes that have yet
     * emerged following some research as being required.
     *
     * Excel has other embedded character escaping requirements, some that emerged
     * from empirical testing, other through research. Firstly, with regards to
     * any embedded speech marks ("), each occurrence should be escaped with
     * another speech mark and the whole field then surrounded with speech marks.
     * Thus if a field holds <em>"Hello" he said</em> then it should be modified
     * to appear as <em>"""Hello"" he said"</em>. Furthermore, if the field
     * contains either embedded separator or EOL characters, it should also
     * be surrounded with speech marks. As a result <em>1,400</em> would become
     * <em>"1,400"</em> assuming that the comma is the required field separator.
     * This has one consequence in, if a field contains embedded speech marks
     * and embedded separator characters, checks for both are not required as the
     * additional set of speech marks that should be placed around ay field
     * containing embedded speech marks will also account for the embedded
     * separator.
     *
     * It is worth making one further note with regard to embedded EOL
     * characters. If the data in a worksheet is exported as a CSV file using
     * Excel itself, then the field will be surounded with speech marks. If the
     * resulting CSV file is then re-imports into another worksheet, the EOL
     * character will result in the original simgle field occupying more than
     * one cell. This same 'feature' is replicated in this classes behaviour.
     *
     * @param field An instance of the String class encapsulating the formatted
     *        contents of a cell on an Excel worksheet.
     * @return A String that encapsulates the formatted contents of that
     *         Excel worksheet cell but with any embedded separator, EOL or
     *         speech mark characters correctly escaped.
     */
    private String escapeEmbeddedCharacters(String field) {
        StringBuffer buffer;

        // If the fields contents should be formatted to confrom with Excel's
        // convention....
        if(this.formattingConvention == ExcelToCSV.EXCEL_STYLE_ESCAPING) {

            // Firstly, check if there are any speech marks (") in the field;
            // each occurrence must be escaped with another set of spech marks
            // and then the entire field should be enclosed within another
            // set of speech marks. Thus, "Yes" he said would become
            // """Yes"" he said"
            if(field.contains("\"")) {
                buffer = new StringBuffer(field.replaceAll("\"", "\\\"\\\""));
                buffer.insert(0, "\"");
                buffer.append("\"");
            }
            else {
                // If the field contains either embedded separator or EOL
                // characters, then escape the whole field by surrounding it
                // with speech marks.
                buffer = new StringBuffer(field);
                if((buffer.indexOf(this.separator)) > -1 ||
                        (buffer.indexOf("\n")) > -1) {
                    buffer.insert(0, "\"");
                    buffer.append("\"");
                }
            }
            return(buffer.toString().trim());
        }
        // The only other formatting convention this class obeys is the UNIX one
        // where any occurrence of the field separator or EOL character will
        // be escaped by preceding it with a backslash.
        else {
            if(field.contains(this.separator)) {
                field = field.replaceAll(this.separator, ("\\\\" + this.separator));
            }
            if(field.contains("\n")) {
                field = field.replaceAll("\n", "\\\\\n");
            }
            return(field);
        }
    }

    /**
     * The main() method contains code that demonstrates how to use the class.
     *
     * @param args An array containing zero, one or more elements all of type
     *        String. Each element will encapsulate an argument specified by the
     *        user when running the program from the command prompt.
     */
    public static void main(String[] args) {
        // Check the number of arguments passed to the main method. There
        // must be two, three or four; the name of and path to either the folder
        // containing the Excel files or an individual Excel workbook that is/are
        // to be converted, the name of and path to the folder to which the CSV
        // files should be written, - optionally - the separator character
        // that should be used to separate individual items (fields) on the
        // lines (records) of the CSV file and - again optionally - an integer
        // that idicates whether the CSV file ought to obey Excel's or UNIX
        // convnetions with regard to formatting fields that contain embedded
        // separator, Speech mark or EOL character(s).
        //
        // Note that the names of the CSV files will be derived from those
        // of the Excel file(s). Put simply the .xls or .xlsx extension will be
        // replaced with .csv. Therefore, if the source folder contains files
        // with matching names but different extensions - Test.xls and Test.xlsx
        // for example - then the CSV file generated from one will overwrite
        // that generated from the other.
        ExcelToCSV converter;
        boolean converted = true;
        long startTime = System.currentTimeMillis();
        try {
            converter = new ExcelToCSV();
            if(args.length == 2) {
                // Just the Source File/Folder and Destination Folder were
                // passed to the main method.
                converter.convertExcelToCSV(args[0], args[1]);
            }
            else if(args.length == 3){
                // The Source File/Folder, Destination Folder and Separator
                // were passed to the main method.
                converter.convertExcelToCSV(args[0], args[1], args[2]);
            }
            else if(args.length == 4) {
                // The Source File/Folder, Destination Folder, Separator and
                // Formatting Convnetion were passed to the main method.
                converter.convertExcelToCSV(args[0], args[1],
                        args[2], Integer.parseInt(args[3]));
            }
            else {
                // None or more than four parameters were passed so display
                //a Usage message.
                System.out.println("Usage: java ToCSV [Source File/Folder] " +
                        "[Destination Folder] [Separator] [Formatting Convention]\n" +
                        "\tSource File/Folder\tThis argument should contain the name of and\n" +
                        "\t\t\t\tpath to either a single Excel workbook or a\n" +
                        "\t\t\t\tfolder containing one or more Excel workbooks.\n" +
                        "\tDestination Folder\tThe name of and path to the folder that the\n" +
                        "\t\t\t\tCSV files should be written out into. The\n" +
                        "\t\t\t\tfolder must exist before running the ToCSV\n" +
                        "\t\t\t\tcode as it will not check for or create it.\n" +
                        "\tSeparator\t\tOptional. The character or characters that\n" +
                        "\t\t\t\tshould be used to separate fields in the CSV\n" +
                        "\t\t\t\trecord. If no value is passed then the comma\n" +
                        "\t\t\t\twill be assumed.\n" +
                        "\tFormatting Convention\tOptional. This argument can take one of two\n" +
                        "\t\t\t\tvalues. Passing 0 (zero) will result in a CSV\n" +
                        "\t\t\t\tfile that obeys Excel's formatting conventions\n" +
                        "\t\t\t\twhilst passing 1 (one) will result in a file\n" +
                        "\t\t\t\tthat obeys UNIX formatting conventions. If no\n" +
                        "\t\t\t\tvalue is passed, then the CSV file produced\n" +
                        "\t\t\t\twill obey Excel's formatting conventions.");
                converted = false;
            }
        }
        // It is not wise to have such a wide catch clause - Exception is very
        // close to being at the top of the inheritance hierarchy - though it
        // will suffice for this example as it is really not possible to recover
        // easilly from an exceptional set of circumstances at this point in the
        // program. It should however, ideally be replaced with one or more
        // catch clauses optimised to handle more specific problems.
        catch(Exception ex) {
            System.out.println("Caught an: " + ex.getClass().getName());
            System.out.println("Message: " + ex.getMessage());
            System.out.println("Stacktrace follows:.....");
            ex.printStackTrace(System.out);
            converted = false;
        }

        if (converted) {
            System.out.println("Conversion took " +
                    ((System.currentTimeMillis() - startTime)/1000) + " seconds");
        }
    }

    /**
     * An instance of this class can be used to control the files returned
     * be a call to the listFiles() method when made on an instance of the
     * File class and that object refers to a folder/directory
     */
    class ExcelFilenameFilter implements FilenameFilter {

        /**
         * Determine those files that will be returned by a call to the
         * listFiles() method. In this case, the name of the file must end with
         * either of the following two extension; '.xls' or '.xlsx'. For the
         * future, it is very possible to parameterise this and allow the
         * containing class to pass, for example, an array of Strings to this
         * class on instantiation. Each element in that array could encapsulate
         * a valid file extension - '.xls', '.xlsx', '.xlt', '.xlst', etc. These
         * could then be used to control which files were returned by the call
         * to the listFiles() method.
         *
         * @param file An instance of the File class that encapsulates a handle
         *             referring to the folder/directory that contains the file.
         * @param name An instance of the String class that encapsulates the
         *             name of the file.
         * @return A boolean value that indicates whether the file should be
         *         included in the array retirned by the call to the listFiles()
         *         method. In this case true will be returned if the name of the
         *         file ends with either '.xls' or '.xlsx' and false will be
         *         returned in all other instances.
         */
        @Override
        public boolean accept(File file, String name) {
            return(name.endsWith(".xls") || name.endsWith(".xlsx"));
        }
    }
}
