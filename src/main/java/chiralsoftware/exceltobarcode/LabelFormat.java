package chiralsoftware.exceltobarcode;

import com.itextpdf.text.PageSize;
import com.itextpdf.text.Rectangle;

/**
 *
 * @author hh
 */
public enum LabelFormat {
    
    AVERY5160(PageSize.LETTER, 3,10, 2 + (float)5 / 8 , 1),
    AVERY5163(PageSize.LETTER, 2, 5, 4, 2),
    AVERY5164(PageSize.LETTER, 2, 3, 4, 3 + 1f /3)
    ;
    
    private final Rectangle pageSize;
    private final int columns;
    private final int rows;
    private final float width; // width, in inches, of a single label
    private final float height; // height, in inches, of a single label
    
    LabelFormat(Rectangle pageSize, int columns, int rows, float width, float height) {
        this.pageSize = pageSize;
        this.columns = columns;
        this.rows = rows;
        this.width = width;
        this.height = height;
    }
    
    public int getColumns() { return columns; }
    public int getRows() { return rows; }
    
    float getMarginTop() {
        return (pageSize.getHeight() / 72 - height * rows) / 2;
    }
    float getWidthPercentage() {
        return 100 * width * columns * 72 / pageSize.getWidth();
    }
    float getMarginLeft() {
        return (pageSize.getWidth() / 72 - width * columns) / 2;
    }

    /** Return the width, in inches */
    float getWidth() {
        return width;
    }

    /** Return the height, in inches */
    float getHeight() {
        return height;
    }
    
    public int getLabelsPerPage() {
        return rows * columns;
    }
    
}
