package chiralsoftware.exceltobarcode;

import static chiralsoftware.exceltobarcode.ExcelColumnStatistics.getTrimmedString;
import static chiralsoftware.exceltobarcode.LineType.TEXT;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.Barcode128;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.context.annotation.Scope;
import org.springframework.context.annotation.ScopedProxyMode;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

/**
 * A simple controller to do everything!
 * 
 */
@Controller
@Scope(value="session",proxyMode=ScopedProxyMode.TARGET_CLASS)
public class MainController {
    private static final Logger LOG = Logger.getLogger(MainController.class.getName());
    private static final String xlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    
    @RequestMapping(value = "/",method = RequestMethod.GET)
    public String main(Model model) {
        if(xSSFWorkbook == null) {
            LOG.info("The file is null");
        } else {
            LOG.info("Here is my multipart file FROM THE SESSION: " + xSSFWorkbook);
        }
        model.addAttribute("sheetPreview", sheetPreview);
        model.addAttribute("optionList", optionList);
        model.addAttribute("columnList",  columnList);
        return "/index";
    }
    
    /** Return a preview of the spreadsheet */
    private List<HtmlCell[]> getSheetPreview(XSSFSheet sheet, List<Integer> columnNumbers) {
        final List<HtmlCell[]> result = new ArrayList<>();
        int rowCount = 0;
        for(Row r: sheet) {
            final HtmlCell[] rowArray = new HtmlCell[columnNumbers.size()];
            int colPos = 0;
            for(int c : columnNumbers) {
                rowArray[colPos] = 
                        new HtmlCell(getTrimmedString(r.getCell(c)), 
                                DisplayUtilities.getTdStyle(r.getCell(c), findFont(r.getCell(c))));
                colPos++;
            }
            result.add(rowArray);
            rowCount++;
            if(rowCount > 1000) break;
        }
        return result;
    }

    /** How many columns are used in this spreadsheet */
    private List<Integer> getLiveColumnList() {
        if(xSSFWorkbook == null) {
            LOG.info("no excel spreadsheet has been uploaded");
            return Collections.EMPTY_LIST;
        }
        
        int count = 0;
        final XSSFSheet sheet = xSSFWorkbook.getSheetAt(0);
        if(sheet == null) {
            LOG.info("Couldn't get a sheet at index 0");
            return Collections.EMPTY_LIST;
        }
        final Set<Integer> result = new HashSet<>();
        for (Row r : sheet) {
            for(Cell c : r) {
                if(c.getCellType() == Cell.CELL_TYPE_BLANK) continue;
                result.add(c.getColumnIndex());
            }
            if(count > 1000) break;
        }
        final List<Integer> l = new ArrayList<>(result.size());
        l.addAll(result);
        return Collections.unmodifiableList(l);
    }
    
    private XSSFWorkbook xSSFWorkbook = null;
    private  List<HtmlCell[]> sheetPreview = null;
    private List<Integer> columnList = null;
    private List<LineConfiguration> lineConfigurations = null;
    
    private static final String lineSelectValuePrefix = "Line ";
    static {
        final List<String> l = new ArrayList<>();
        l.add("-");
        for(int i = 1; i <= 10; i++)
            l.add(lineSelectValuePrefix + i);
        optionList = Collections.unmodifiableList(l);
    }
    
    private static final List<String> optionList;
    
    @RequestMapping(value = "/",method = RequestMethod.POST)
    public String excelUpload(@RequestParam("fileToUpload") MultipartFile file, 
            Model model, RedirectAttributes redirectAttributes) throws IOException {
        LOG.info("Here I am and here is the multipart:");
        if(file == null) {
            LOG.info("OH no! file was null!");
            return "/index";
        }
        LOG.info("File: " + file.getName() + " / " + file.getSize());
        if(! xlsxContentType.equalsIgnoreCase(file.getContentType())) {
            redirectAttributes.addAttribute("message", "File type was wrong - it should be .xlsx");
            LOG.info("Content type was wrong; expected " + xlsxContentType + " but got: " + file.getContentType());
            return "redirect:/index.html";
        }
//        if(! file.getOriginalFilename().endsWith(".xlsx")) {
//            LOG.info("Problem: file name was wrong extension: " + file.getOriginalFilename());
//            redirectAttributes.addAttribute("message", "File type was wrong - it should be .xlsx");
//            return "redirect:/";
//        }
        LOG.info("about to get input stream");
        // fixme - we should allow access to more than one sheet
        xSSFWorkbook =  new XSSFWorkbook(file.getInputStream());
        // This is a list of integers corresponding to the live columns
        columnList = getLiveColumnList();
        model.addAttribute("columnList",  columnList);
        sheetPreview = getSheetPreview(xSSFWorkbook.getSheetAt(0),columnList);
        model.addAttribute("sheetPreview", sheetPreview);
        model.addAttribute("optionList", optionList);
        return "/index";
    }
    
    // <editor-fold desc="show a blank document" defaultstate="collapsed">
    /** Shows a blank document, in case of a problem in generating the PDF */
    private byte[] showBlank() throws DocumentException {
        final Document document = new Document(PageSize.LETTER); // FIXME - get PageSize from label definition
        final ByteArrayOutputStream baos = new ByteArrayOutputStream();
        
        final PdfWriter writer = PdfWriter.getInstance(document, baos);
        document.open();
        document.add(new Paragraph("No data have been uploaded.  The time is: " + new Date()));
        
        final Barcode128 code128 = new Barcode128();
        code128.setGenerateChecksum(true);
        code128.setCode(new Date().toString());
        
        document.add(code128.createImageWithBarcode(writer.getDirectContent(), null, null));
        document.close();
        return baos.toByteArray();
    }
    // </editor-fold>
    
    private Font findFont(org.apache.poi.ss.usermodel.Cell cell) {
        if(cell == null) return new Font(Font.FontFamily.HELVETICA, 10);
        if(xSSFWorkbook == null)
            throw new NullPointerException("workbook was null!");
        final org.apache.poi.ss.usermodel.Font f = xSSFWorkbook.getFontAt(cell.getCellStyle().getFontIndex());
        

        final Font.FontFamily family;
        final String fontName = f.getFontName().toLowerCase();
        if(fontName.contains("times")) family = Font.FontFamily.TIMES_ROMAN;
        else if(fontName.contains("courier")) family = Font.FontFamily.COURIER;
        else if(fontName.contains("symbol")) family = Font.FontFamily.SYMBOL;
        else family = Font.FontFamily.HELVETICA;
        
        //final Font result = new Font(family, size);
        // the font height value is 20 * the font size in points
        final Font result = new Font(family, (float) f.getFontHeight() / 20);
        
        if(f.getBold()) result.setStyle(Font.BOLD);
        if(f.getItalic()) result.setStyle(Font.ITALIC);
        return result;
    }
    
    /** What is the font size as a ratio of the total line height.  This needs to leave
     * space for descenders
     */
    private static final float fontSizeRatio = 0.75f;
    
    /** Show one label by taking one row and printing it on the document */
    private void showOneLabel(Row r, PdfPTable t, PdfWriter writer, LabelFormat labelFormat) throws DocumentException {
        if(t == null) throw new NullPointerException("table parameter was null");
        if(r == null) {
            // there is a missing row
            // so insert a blank cell and move on
            final PdfPCell cell = new PdfPCell();
            cell.setFixedHeight(labelFormat.getHeight() * 72);
            cell.setPadding(0.1f * 72); // a tenth of an inch
            cell.setHorizontalAlignment(PdfPCell.ALIGN_CENTER);
            cell.addElement(new Paragraph(" ")); // FIXME - make this more smart
            t.addCell(cell);
            return;
        }
        if(lineConfigurations == null)
            throw new NullPointerException("Line configurations was null!");
        final PdfPCell cell = new PdfPCell();
        cell.setFixedHeight(labelFormat.getHeight() * 72);
        cell.setPadding(0.1f * 72); // a tenth of an inch
        cell.setHorizontalAlignment(PdfPCell.ALIGN_CENTER);
        cell.setBorder(Rectangle.NO_BORDER);
        // calculate barcode height
        final float barcodeHeight = cell.getFixedHeight() / lineConfigurations.size();
        for(LineConfiguration lineConfiguration : lineConfigurations) {
            LOG.info("Showing row: " + r.getRowNum() + ", line config: " + lineConfiguration);
            if(lineConfiguration.getLineType() == TEXT) {
                final org.apache.poi.ss.usermodel.Cell excelCell = 
                        r.getCell(lineConfiguration.getColumnNumber());
                final Cell myCell = r.getCell(lineConfiguration.getColumnNumber());
                if(myCell == null) {
                    cell.addElement(new Paragraph(" "));
                } else {
                    final Paragraph p = 
                            new Paragraph(getTrimmedString(myCell), findFont(excelCell));
                    p.setLeading(0.8f * fontSizeRatio * barcodeHeight);
//                    p.setLeading(10);
                    if(excelCell.getCellStyle().getAlignment() == CellStyle.ALIGN_LEFT) 
                        p.setAlignment(Paragraph.ALIGN_LEFT);
                    else if(excelCell.getCellStyle().getAlignment() == CellStyle.ALIGN_CENTER) 
                        p.setAlignment(Paragraph.ALIGN_CENTER);
                    else if(excelCell.getCellStyle().getAlignment() == CellStyle.ALIGN_RIGHT) 
                        p.setAlignment(Paragraph.ALIGN_RIGHT);
                    p.setSpacingAfter(barcodeHeight * 0.2f);
                    cell.addElement(p);
                }
            } else {
                final Barcode128 code128 = new Barcode128();
                code128.setGenerateChecksum(true);
                final Cell myCell = r.getCell(lineConfiguration.getColumnNumber());
                if(myCell == null) code128.setCode(" ");
                else code128.setCode(getTrimmedString(myCell));
                final Image image = code128.createImageWithBarcode(writer.getDirectContent(), null, null);
                image.scaleToFit(labelFormat.getWidth() * 72, 0.9f * barcodeHeight);
                image.setAlignment(Image.ALIGN_CENTER);
                cell.addElement(image);
            }
            // set cell height:
            // http://itextpdf.com/examples/iia.php?id=81
        }
        t.addCell(cell);
    }
    
    private static LineType findLineType(String s) {
        if(s == null) return TEXT;
        if(s.equalsIgnoreCase("barcode")) return LineType.BARCODE;
        return TEXT;
    }
    
    /** Take the request parameter map, and also look at the live column list,
     * and create a set of line definitions
     * @param allRequestParams 
     */
    private void createLineTypes(Map<String,String> allRequestParams) {
//        LOG.info("This is the sorted key list: " + allRequestParams.keySet());
        final List<LineConfiguration> l = new ArrayList<>();
        for(String s : allRequestParams.keySet()) {
            if(s.startsWith("column_") && (! s.startsWith("column_t"))) {
                int columnNumber = Integer.parseInt(s.substring("column_".length()));
                if(! allRequestParams.get(s).startsWith(lineSelectValuePrefix)) continue;
                final int lineNumber = Integer.parseInt(allRequestParams.get(s).substring(lineSelectValuePrefix.length()));
//                LOG.info("Column number: " + columnNumber + " should go to line number: " + lineNumber);
                if(! columnList.contains(columnNumber)) {
                    LOG.info("The column list: " + columnList + " "
                            + "doesn't contain the supplied column number: " + columnNumber + " so skipping");
                    continue;
                }
                final LineConfiguration lineConfiguration = new LineConfiguration(columnNumber, lineNumber,
                        findLineType(allRequestParams.get("column_type_" + columnNumber)));
                l.add(lineConfiguration);
            }
        }
        // now sort this list according to the order of line numbers
        Collections.sort(l, LineConfiguration.sortByLineNumber);
        this.lineConfigurations = Collections.unmodifiableList(l);
//        LOG.info("Here is the line configurations list: " + lineConfigurations);
    }
    
    private PdfPTable createTable(PdfWriter writer, LabelFormat labelFormat, int firstRow, int lastRow) throws DocumentException {
        if(firstRow >= lastRow) {
            LOG.info("First row = " + firstRow + ", last row = " + lastRow + " so returning null");
            return null;
        }
        if(lastRow - firstRow > labelFormat.getColumns() * labelFormat.getRows()) {
            LOG.info("There were too many labels requested.  "
                    + "You wanted: " + (lastRow - firstRow) + " labels, "
                    + "but this label format only allows: " + labelFormat.getColumns() * labelFormat.getRows());
            return null;
        }
        final XSSFSheet sheet = xSSFWorkbook.getSheetAt(0);
        if(lastRow > sheet.getLastRowNum()) {
            LOG.info("last row was larger that the last row number of this sheet: " + sheet.getLastRowNum());
            lastRow = sheet.getLastRowNum();
        }
        if(firstRow >= lastRow) {
            LOG.info("after adjusting the last row, First row = " + firstRow + ", last row = " + lastRow + " so returning null");
            return null;
        }
        
        final PdfPTable table = new PdfPTable(labelFormat.getColumns());
        table.getDefaultCell().setBorder(PdfPCell.NO_BORDER);
        table.setWidthPercentage(labelFormat.getWidthPercentage());   
        LOG.info("The width percentage i found is; " + labelFormat.getWidthPercentage());
        for(int i = firstRow; i <= lastRow; i++) {
            LOG.info("Showing row: " + i);
            showOneLabel(sheet.getRow(i), table, writer, labelFormat);
        }
        if(lastRow == sheet.getLastRowNum()) {
            // we are at the end
            //check to make sure that the last row is filled in
            if(lastRow % labelFormat.getColumns() != 0) {
                for(int x = 0; x < lastRow % labelFormat.getColumns(); x++) {
                    final PdfPCell cell = new PdfPCell();
                    cell.setFixedHeight(labelFormat.getHeight() * 72);
                    cell.addElement(new Paragraph("  "));
                    table.addCell(cell);
                }
                    
            }
        }
        return table;
    }
    
    @RequestMapping(value = "/export.pdf")
    public ResponseEntity<byte[]> generatePdf(@RequestParam Map<String,String> allRequestParams) 
            throws DocumentException {
        final LabelFormat labelFormat = allRequestParams.containsKey("labelFormatString") ?
                LabelFormat.valueOf(allRequestParams.get("labelFormatString")) :
                LabelFormat.AVERY5160;
        
        createLineTypes(allRequestParams);

        if(xSSFWorkbook == null) {
            LOG.info("The workbook is null so this wouldn't work really");
            return new ResponseEntity<>(showBlank(), HttpStatus.OK);
        }
        if(allRequestParams == null) {
            LOG.info("the allRequestParams param is null so this wouldn't work really");
            return new ResponseEntity<>(showBlank(), HttpStatus.OK);
        }
        
        // we create a new document with zero left/right margins
        // we calculate the top and bottom margin
        final float topMargin = (PageSize.LETTER.getHeight() - labelFormat.getRows() * labelFormat.getHeight() * 72) / 2;
        final Document document = new Document(PageSize.LETTER, 0,0,topMargin,topMargin);
        final ByteArrayOutputStream baos = new ByteArrayOutputStream();
        
        final PdfWriter writer = PdfWriter.getInstance(document, baos);
        
        document.open();
        
        final XSSFSheet sheet = xSSFWorkbook.getSheetAt(0);
        final int rowCount = sheet.getLastRowNum();
        LOG.info("With: " + rowCount + " rows, "
                + "and " + labelFormat.getRows() * labelFormat.getColumns() + " labels per page, "
                + "we need: " + (1 + rowCount / (labelFormat.getRows() * labelFormat.getColumns())) + " pages");
        for(int i = 0 ; i <=  rowCount / labelFormat.getLabelsPerPage(); i++) {
            LOG.info("Showing page: " + i);
            int firstRow = i * labelFormat.getLabelsPerPage();
            int lastRow =  firstRow + labelFormat.getLabelsPerPage();
            if(lastRow > rowCount) lastRow = rowCount;
            LOG.info("At i = " + i + ", we need to show rows " + firstRow + " to " + lastRow);
            if(lastRow > firstRow) {
                final PdfPTable t = createTable(writer, labelFormat, firstRow, lastRow);
                document.add(t);
                t.setComplete(true);
                LOG.info("i = " + i + ", added the table and adding a new page");
                document.newPage();
            }
        } 
        
        document.close();
        
        final HttpHeaders httpHeaders = new HttpHeaders();
        httpHeaders.setContentType(MediaType.APPLICATION_PDF);
        final ResponseEntity<byte[]> result = new ResponseEntity<>(baos.toByteArray(), 
                httpHeaders, HttpStatus.OK);
        return result;
    }
    
    
}
