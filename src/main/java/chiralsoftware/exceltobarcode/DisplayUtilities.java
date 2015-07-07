package chiralsoftware.exceltobarcode;

import com.itextpdf.text.Font;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 *
 * @author hh
 */
public final class DisplayUtilities {
    private DisplayUtilities() {
        throw new RuntimeException("Don't instantiate this!");
    }
    
        /** Find the appropriate CSS style for a given cell.
     * This finds out the font family, bold, italic, alignment,
     * and size.
     * 
     * @param c
     * @param font
     * @return 
     */
    public static final String getTdStyle(Cell c, Font font) {
        if(c == null) return "";
        if(font == null) return "";
        
        final StringBuilder result = new StringBuilder();
        final CellStyle cellStyle = c.getCellStyle();
        if(cellStyle.getAlignment() == CellStyle.ALIGN_LEFT) 
            result.append("text-align: left; ");
        else if(cellStyle.getAlignment() == CellStyle.ALIGN_CENTER) 
            result.append("text-align: center; ");
        else if(cellStyle.getAlignment() == CellStyle.ALIGN_RIGHT) 
            result.append("text-align: right; ");
        else result.append("text-align: left; ");
        
        if(font.getFamily() == Font.FontFamily.COURIER) 
            result.append("font-family: Courier, monospace; ");
        else if(font.getFamily() == Font.FontFamily.HELVETICA) 
            result.append("font-family: sans-serif; ");
        else if(font.getFamily() == Font.FontFamily.SYMBOL) 
            result.append("font-family: sans-serif; ");
        else if(font.getFamily() == Font.FontFamily.TIMES_ROMAN) 
            result.append("font-family: Times, serif; ");
        
        result.append("font-size: ").append(Math.round(font.getSize())).append("pt;");
        
        if(font.isBold()) 
            result.append("font-weight: bold; ");
        if(font.isItalic())
            result.append("font-style: italic; ");
        
        return result.toString();

    }

    
}
