package chiralsoftware.exceltobarcode;

import com.itextpdf.text.Font;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

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
        if(null == cellStyle.getAlignmentEnum()) 
            result.append("text-align: left; ");
        else switch (cellStyle.getAlignmentEnum()) {
            case LEFT:
                result.append("text-align: left; ");
                break;
            case CENTER:
                result.append("text-align: center; ");
                break;
            case RIGHT:
                result.append("text-align: right; ");
                break;
            default:
                result.append("text-align: left; ");
                break;
        }
        
        if(null != font.getFamily()) 
            switch (font.getFamily()) {
            case COURIER:
                result.append("font-family: Courier, monospace; ");
                break;
            case HELVETICA:
                result.append("font-family: sans-serif; ");
                break;
            case SYMBOL:
                result.append("font-family: sans-serif; ");
                break;
            case TIMES_ROMAN:
                result.append("font-family: Times, serif; ");
                break;
            default:
                break;
        }
        
        result.append("font-size: ").append(Math.round(font.getSize())).append("pt;");
        
        if(font.isBold()) 
            result.append("font-weight: bold; ");
        if(font.isItalic())
            result.append("font-style: italic; ");
        
        return result.toString();

    }

    
}
