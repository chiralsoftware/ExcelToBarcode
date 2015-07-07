package chiralsoftware.exceltobarcode;

/**
 *
 */
public final class HtmlCell {
    
    HtmlCell(String text, String style) {
        if(text == null) 
            throw new NullPointerException("The text param was null!");
        if(style == null)
            throw new NullPointerException("The style param was null!");
        
        this.text = text;
        this.style = style;
    }
    
    private final String text;
    private final String style;

    public String getText() {
        return text;
    }

    public String getStyle() {
        return style;
    }

    @Override
    public String toString() {
        return "HtmlCell{" + "text=" + text + ", style=" + style + '}';
    }
    
}
