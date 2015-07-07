package chiralsoftware.exceltobarcode;

import au.com.bytecode.opencsv.CSVReader;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.Cell;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BOOLEAN;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_ERROR;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING;

/**
 *
 * @author hh
 */
final class ExcelColumnStatistics {

    private static final Logger LOG = Logger.getLogger(ExcelColumnStatistics.class.getName());

    private static final String EMAIL_PATTERN
            = "^[_A-Za-z0-9-\\+]+(\\.[_A-Za-z0-9-]+)*@"
            + "[A-Za-z0-9-]+(\\.[A-Za-z0-9]+)*(\\.[A-Za-z]{2,})$";
    private static final Pattern pattern = Pattern.compile(EMAIL_PATTERN);

    /**
     * Test if a string is a likely phone number or not. String must be
     * non-null, non-empty, stripped
     *
     * @param s string to test
     * @return true if this is a likely phone number
     */
    private static boolean isPhone(String s) {
        if (s.length() > 25) return false;
        final String digitsOnly = s.replaceAll("\\D", "");
        if (digitsOnly.length() < 10) return false;
        if (digitsOnly.length() > 17) return false;
        return true;
    }
    private static final Set<String> stateSet;
    private static final Set<String> citySet;
    private static final Set<String> lastNameSet;

    static {
        final Set<String> s = new HashSet<>();
        final String[] stateArray = {
            "al", "alabama",
            "ak", "alaska",
            "az", "arizona",
            "ar", "arkansas",
            "ca", "california",
            "co", "colorado",
            "ct", "connecticut",
            "de", "delaware",
            "fl", "florida",
            "ga", "georgia",
            "hi", "hawaii",
            "id", "idaho",
            "il", "illinois",
            "in", "indiana",
            "ia", "iowa",
            "ks", "kansas",
            "ky", "kentucky",
            "la", "louisiana",
            "me", "maine",
            "md", "maryland",
            "ma", "massachusetts",
            "mi", "michigan",
            "mn", "minnesota",
            "ms", "mississippi",
            "mo", "missouri",
            "mt", "montana",
            "ne", "nebraska",
            "nv", "nevada",
            "nh", "new hampshire",
            "nj", "new jersey",
            "nm", "new mexico",
            "ny", "new york",
            "nc", "north carolina",
            "nd", "north dakota",
            "oh", "ohio",
            "ok", "oklahoma",
            "or", "oregon",
            "pa", "pennsylvania",
            "ri", "rhode island",
            "sc", "south carolina",
            "sd", "south dakota",
            "tn", "tennessee",
            "tx", "texas",
            "ut", "utah",
            "vt", "vermont",
            "va", "virginia",
            "wa", "washington",
            "wv", "west virginia",
            "wi", "wisconsin",
            "wy", "wyoming",
            "dc", "washington dc",
            // canada
            "ab", "alberta",
            "bc", "british columbia",
            "mb", "manitoba",
            "nb", "new brunswick",
            "nl", "newfoundland and labrador",
            "ns", "nova scotia",
            "on", "ontario",
            "pe", "prince edward island",
            "qc", "quebec",
            "sk", "saskatchewan",
            // australia
            "australian capital territory", "act",
            "jervis bay territory", "jbt",
            "new south wales", "nsw",
            "norfolk island", "nf",
            "northern territory", "nt",
            "queensland", "qld",
            "south australia", "sa",
            "tasmania", "tas",
            "victoria", "vic",
            "western australia", "wa",};
        for(String stateString : stateArray) s.add(stateString.replaceAll("[^a-z]", ""));
        s.addAll(Arrays.asList(stateArray));
        stateSet = Collections.unmodifiableSet(s);
        // load the USA cities from CSV resource
        final String usCitiesResource = "/cities.csv";
        InputStream is = ExcelColumnStatistics.class.getResourceAsStream(usCitiesResource);
        if(is == null) throw new NullPointerException("Couldn't find a resource for: " + usCitiesResource);
        CSVReader reader = new CSVReader(new InputStreamReader(is));
        final Set<String> cs = new HashSet<>();
        String[] nextLine;
        try {
            while ((nextLine = reader.readNext()) != null) {
                // nextLine[] is an array of values from the line
                cs.add(nextLine[2].toLowerCase().replaceAll("[^a-z]", ""));
            }
            is.close();
        } catch (IOException ex) {
            LOG.log(Level.SEVERE, "couldn't read: " + usCitiesResource, ex);
        }
        // load Australian cities from CSV resource
        final String australiaCitiesResource = "/australiantownslist.csv";
        is = ExcelColumnStatistics.class.getResourceAsStream(australiaCitiesResource);
        reader = new CSVReader(new InputStreamReader(is));
        try {
            reader.readNext(); // first line is title
            while ((nextLine = reader.readNext()) != null) {
                // nextLine[] is an array of values from the line
                cs.add(nextLine[0].toLowerCase().replaceAll("[^a-z]", ""));
            }
            is.close();
        } catch (IOException ex) {
            LOG.log(Level.SEVERE, "couldn't read: " + usCitiesResource, ex);
        }
        citySet = Collections.unmodifiableSet(cs);

        // get the set of likely last name
        final String lastNamesResource = "/last-names.txt";
        is = ExcelColumnStatistics.class.getResourceAsStream(lastNamesResource);
        final BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(is));
        String oneLine;
        final Set<String> ns = new HashSet();
        try {
            while ((oneLine = bufferedReader.readLine()) != null)
                ns.add(oneLine.trim().toLowerCase());
            is.close();
        } catch (IOException ex) {
            LOG.log(Level.SEVERE, "couldn't read: " + usCitiesResource, ex);
        }
        lastNameSet = Collections.unmodifiableSet(ns);
    }

    private int rowCount = 0;
    private int emails = 0;
    private int phones = 0;
    private int states = 0;
    private int cities = 0;
    private int names = 0;
    private int blank;

    /** Null-safe string trim */
    private static String trim(String s) {
        if(s == null) return null;
        return s.trim();
    }

    static final String getTrimmedString(Cell c) {
        if(c == null) return "";
        switch(c.getCellType()) {
            case CELL_TYPE_BLANK: return "";
            case CELL_TYPE_BOOLEAN: return Boolean.toString(c.getBooleanCellValue()); 
            case CELL_TYPE_ERROR: return "";
            case CELL_TYPE_FORMULA: return "";
            case CELL_TYPE_NUMERIC: return Long.toString(Math.round(c.getNumericCellValue()));
            case CELL_TYPE_STRING: return trim(c.getStringCellValue());
        }
        LOG.warning("Unknown cell type: " + c.getCellType());
        return "";
    }
    
    void update(Cell cell) {
        rowCount++;
        String string = getTrimmedString(cell);
        if (string == null) {
            blank++;
            return;
        }
        string = string.trim();
        if (string.isEmpty()) {
            blank++;
            return;
        }
        if (pattern.matcher(string).matches()) {
            emails++;
            return;
        }
        if (isPhone(string)) {
            phones++;
            return;
        }
        string = string.toLowerCase();
        final String noBlankString = string.replaceAll("[^a-z]", "");
        if (stateSet.contains(noBlankString)) {
            states++;
            return;
        }
        if (citySet.contains(noBlankString)) {
            cities++;
        }
        // now we tokenize the string, and see if it might be a name
        final String[] splitString = string.split("\\s");
        if(splitString.length < 4) 
            for(String possibleName : splitString) {
                if(lastNameSet.contains(possibleName)) {
                    names++;
                    break;
                }
            }

    }

    private static final Option noneOption = new Option("none", "---");
    private static final Option nameOption = new  Option("name", "name");
    private static final Option organizationOption = new  Option("organization", "company");
    private static final Option titleOption = new  Option("title", "title");
    private static final Option phoneOption = new  Option("phone", "phone");
    private static final Option emailOption = new  Option("email", "email");
    private static final Option notesOption = new  Option("notes", "notes");
    private static final Option addressLine1Option = new  Option("addressLine1", "address 1");
    private static final Option addressLine2Option = new  Option("addressLine2", "address 2");
    private static final Option cityOption = new  Option("city", "city");
    private static final Option provinceOption = new  Option("province", "state");
    private static final Option postalCodeOption = new  Option("postalCode", "zip code");
    private static final Option countryOption = new  Option("country", "country");
    private static final Option categoryOption = new  Option("category", "category");
    
    private static final List<Option> anyColumn; // for a column that we can't guess what it is
    private static final List<Option> blankColumn; // nothing to import
    private static final List<Option> stateColumn;
    private static final List<Option> cityColumn;
    private static final List<Option> phoneColumn;
    private static final List<Option> emailColumn;
    private static final List<Option> nameColumn;

    static {
        List<Option> l = new ArrayList<>();
        l.add(noneOption.select());
        l.add(nameOption);
        l.add(organizationOption);
        l.add(titleOption);
        l.add(phoneOption);
        l.add(emailOption);
        l.add(notesOption);
        l.add(addressLine1Option);
        l.add(addressLine2Option);
        l.add(cityOption);
        l.add(provinceOption);
        l.add(postalCodeOption);
        l.add(countryOption);
        l.add(categoryOption);
        anyColumn = Collections.unmodifiableList(l);
        
        l = new ArrayList<>();
        l.add(noneOption.select());
        blankColumn = Collections.unmodifiableList(l);

        l = new ArrayList<>();
        l.add(noneOption);
        l.add(provinceOption.select());
        stateColumn = Collections.unmodifiableList(l);
        
        l = new ArrayList<>();
        l.add(noneOption);
        l.add(cityOption.select());
        cityColumn = Collections.unmodifiableList(l);

        l = new ArrayList<>();
        l.add(noneOption);
        l.add(nameOption.select());
        nameColumn = Collections.unmodifiableList(l);
        
        l = new ArrayList<>();
        l.add(noneOption);
        l.add(phoneOption.select());
        phoneColumn = Collections.unmodifiableList(l);
        l = new ArrayList<>();
        l.add(noneOption);
        l.add(emailOption.select());
        emailColumn = Collections.unmodifiableList(l);
    }

    @Override
    public String toString() {
        return "ExcelColumnStatistics{" + "rowCount=" + rowCount + ", emails=" + emails + ", phones=" + phones + ", states=" + states + ", cities=" + cities + ", blank=" + blank + '}';
    }

    public List<Option> columnGuesses() {
        final List<Option> result = new ArrayList<>();
        result.add(new Option("none", "------", false));
        
        if (((float) emails / rowCount) > 0.8) return emailColumn;

        if (((float) states / rowCount) > 0.8) return stateColumn;
        
        if (((float) cities / rowCount) > 0.8) return cityColumn;
        
        if (((float) phones / rowCount) > 0.8) return phoneColumn;

        if (((float) names / rowCount) > 0.9) return nameColumn;
        
        if (((float) blank / rowCount) > 0.99) return blankColumn;
        
        // at this point it could be anything, except if there are no emails or phones, we can exclude those
        final List<Option> remaining = new ArrayList(anyColumn);
        if (((float) emails / rowCount) < 0.2) remaining.remove(emailOption);
        if (((float) phones / rowCount) < 0.2) remaining.remove(phoneOption);
        // we should be smart, and if there is already a state or city column, remove that from the choices

        return Collections.unmodifiableList(remaining);
    }
    
    private static final class Option {
        private final String value;
        private final String name;
        private final boolean selected;
        private final Option selectedOption;
        
        private Option(String value, String name) {
            this.value = value;
            this.name = name;
            this.selected = false;
            this.selectedOption = new Option(value, name, true);
        }
        
        @SuppressWarnings("LeakingThisInConstructor")
        private Option(String value, String name, boolean selected) {
            this.value = value;
            this.name = name;
            this.selected = selected;
            this.selectedOption = this;
        }
        
        /** Create a new option which is a copy of this, but it is selected */
        public Option select() {
            return selectedOption;
        }

        public String getValue() {
            return value;
        }

        public String getName() {
            return name;
        }

        public boolean isSelected() {
            return selected;
        }

        @Override
        public String toString() {
            return "Option{" + "value=" + value + ", name=" + name + ", selected=" + selected + '}';
        }

        @Override
        public int hashCode() {
            return Objects.hashCode(value);
        }

        @Override
        public boolean equals(Object obj) {
            if (obj == null)
                return false;
            if (getClass() != obj.getClass())
                return false;
            final Option other = (Option) obj;
            if (!Objects.equals(this.value, other.value))
                return false;
            return true;
        }
        
    }
}
