package chiralsoftware.exceltobarcode;

import java.util.Comparator;

/**
 * Every line has a configuration which includes which column number it
 * corresponds to, and how it is displayed: text or barcode. Later on we will
 * add options such as which type of barcode, text font, text alignment, etc.
 */
final class LineConfiguration {

    private final int columnNumber;
    private final int lineNumber;

    private final LineType lineType;

    public LineConfiguration(int columnNumber, int lineNumber, LineType lineType) {
        if (lineType == null) {
            throw new NullPointerException("Line type was null");
        }

        this.columnNumber = columnNumber;
        this.lineType = lineType;
        this.lineNumber = lineNumber;
    }

    public int getColumnNumber() {
        return columnNumber;
    }

    public int getLineNumber() {
        return lineNumber;
    }

    public LineType getLineType() {
        return lineType;
    }

    @Override
    public String toString() {
        return "LineConfiguration{" + "columnNumber=" + columnNumber + ", "
                + "lineNumber=" + lineNumber + ", lineType=" + lineType + '}';
    }

    static final Comparator<LineConfiguration> sortByLineNumber
            = new Comparator<LineConfiguration>() {

                @Override
                public int compare(LineConfiguration arg0, LineConfiguration arg1) {
                    return arg0.getLineNumber() - arg1.getLineNumber();
                }
            };

}
